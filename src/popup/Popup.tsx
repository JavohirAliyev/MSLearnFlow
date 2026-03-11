import React, { useEffect, useState, useCallback, useRef } from 'react';
import {
  detectContentType,
  fetchCourseStructure,
  fetchLearningPathStructure,
  fetchModuleStructure,
  moduleUrlFromUnitUrl,
  CourseNode,
  LearningPathNode,
  ModuleNode,
  UnitNode,
  ContentType,
} from '../services/scraper';
import { Download, ChevronRight, ChevronDown, Loader2 } from 'lucide-react';
import './popup.css';

// ─── Job state (written by offscreen worker, read by popup) ─────────────────

interface PendingJob {
  units: { title: string; url: string }[];
  docTitle: string;
  filename: string;
}

export interface JobState {
  status: 'idle' | 'fetching' | 'rendering' | 'done' | 'error';
  pct: number;
  label: string;
  docTitle: string;
  totalUnits: number;
  error?: string;
}

// ─── State types ──────────────────────────────────────────────────────────────

type RootData =
  | { type: 'course'; node: CourseNode }
  | { type: 'learning-path'; node: LearningPathNode }
  | { type: 'module'; node: ModuleNode }
  | { type: 'unit'; unitUrl: string; moduleNode: ModuleNode | null };

// ─── Pure tree-update helpers (immutable) ────────────────────────────────────

function setLPInCourse(
  course: CourseNode,
  lpUrl: string,
  update: Partial<LearningPathNode>
): CourseNode {
  return {
    ...course,
    learningPaths: course.learningPaths.map((lp) =>
      lp.url === lpUrl ? { ...lp, ...update } : lp
    ),
  };
}

function setModuleInLP(
  lp: LearningPathNode,
  modUrl: string,
  update: Partial<ModuleNode>
): LearningPathNode {
  return {
    ...lp,
    modules: lp.modules.map((m) => (m.url === modUrl ? { ...m, ...update } : m)),
  };
}

function setModuleInCourse(
  course: CourseNode,
  lpUrl: string,
  modUrl: string,
  update: Partial<ModuleNode>
): CourseNode {
  return {
    ...course,
    learningPaths: course.learningPaths.map((lp) =>
      lp.url === lpUrl ? setModuleInLP(lp, modUrl, update) : lp
    ),
  };
}

// ─── Collect all loaded unit URLs from a subtree ──────────────────────────────

function unitsOfModule(m: ModuleNode): string[] {
  return m.units.map((u) => u.url);
}

function unitsOfLP(lp: LearningPathNode): string[] {
  return lp.modules.flatMap((m) => unitsOfModule(m));
}

function allUnitsOfRoot(root: RootData | null): string[] {
  if (!root) return [];
  if (root.type === 'course') return root.node.learningPaths.flatMap((lp) => unitsOfLP(lp));
  if (root.type === 'learning-path') return unitsOfLP(root.node);
  if (root.type === 'module') return unitsOfModule(root.node);
  if (root.type === 'unit')
    return root.moduleNode ? unitsOfModule(root.moduleNode) : root.unitUrl ? [root.unitUrl] : [];
  return [];
}

// ─── Check-state helpers ──────────────────────────────────────────────────────

type CheckState = 'all' | 'some' | 'none';

function checkStateFor(urls: string[], selected: Set<string>): CheckState {
  if (urls.length === 0) return 'none';
  const count = urls.filter((u) => selected.has(u)).length;
  if (count === 0) return 'none';
  if (count === urls.length) return 'all';
  return 'some';
}

// ─── Indeterminate checkbox ───────────────────────────────────────────────────

interface TriCheckboxProps {
  state: CheckState;
  onChange: () => void;
  disabled?: boolean;
}

const TriCheckbox: React.FC<TriCheckboxProps> = ({ state, onChange, disabled }) => {
  const ref = useRef<HTMLInputElement>(null);
  useEffect(() => {
    if (ref.current) ref.current.indeterminate = state === 'some';
  }, [state]);
  return (
    <input
      ref={ref}
      className="checkbox"
      type="checkbox"
      checked={state === 'all'}
      onChange={onChange}
      disabled={disabled}
    />
  );
};

// ─── Type badge ───────────────────────────────────────────────────────────────

const TYPE_LABELS: Record<string, string> = {
  course: 'Course',
  'learning-path': 'Learning Path',
  module: 'Module',
  unit: 'Unit',
};

const BadgeTag: React.FC<{ type: ContentType | string }> = ({ type }) =>
  TYPE_LABELS[type] ? (
    <span className={`badge badge--${type}`}>{TYPE_LABELS[type]}</span>
  ) : null;

// ─── Session-storage persistence ─────────────────────────────────────────────

interface PersistedState {
  url: string;
  rootData: RootData;
  selected: string[];
  expanded: string[];
}

async function loadSavedState(url: string): Promise<PersistedState | null> {
  try {
    const res = await chrome.storage.session.get('mlf');
    const s = res?.mlf as PersistedState | undefined;
    return s?.url === url ? s : null;
  } catch {
    return null;
  }
}

async function saveState(
  url: string,
  root: RootData,
  sel: Set<string>,
  exp: Set<string>,
): Promise<void> {
  try {
    const state: PersistedState = { url, rootData: root, selected: [...sel], expanded: [...exp] };
    await chrome.storage.session.set({ mlf: state });
  } catch { /* not critical */ }
}

// ─── Main Popup ───────────────────────────────────────────────────────────────

export const Popup: React.FC = () => {
  const [currentUrl, setCurrentUrl] = useState<string | null>(null);
  const [rootData, setRootData] = useState<RootData | null>(null);
  const rootDataRef = useRef<RootData | null>(null);
  const [expanded, setExpanded] = useState<Set<string>>(new Set());
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [loadingIds, setLoadingIds] = useState<Set<string>>(new Set());
  const [loading, setLoading] = useState(false);
  const [jobState, setJobState] = useState<JobState | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  // ── Init: restore from session or do a fresh fetch ──────────────────────────

  useEffect(() => {
    let cancelled = false;

    (async () => {
      // 1. Find active tab URL
      const tabs = await new Promise<chrome.tabs.Tab[]>((r) =>
        chrome.tabs.query({ active: true, lastFocusedWindow: true }, r)
      );
      const url = tabs[0]?.url ?? null;
      if (!url || cancelled) return;
      setCurrentUrl(url);

      // 2. Restore any in-progress or completed job from this session
      const { mlfJob } = await chrome.storage.session.get('mlfJob');
      if (mlfJob && (mlfJob as JobState).status !== 'idle' && !cancelled) {
        setJobState(mlfJob as JobState);
      }

      // 3. Try to restore previously saved tree state for this URL
      const saved = await loadSavedState(url);
      if (saved && !cancelled) {
        setRootData(saved.rootData);
        setSelected(new Set(saved.selected));
        setExpanded(new Set(saved.expanded));
        return; // skip fresh network load
      }

      if (cancelled) return;

      // 3. Fresh load
      setErrorMsg(null);
      setLoading(true);
      try {
        const ct = detectContentType(url);
        if (cancelled) return;

        if (ct === 'course') {
          const node = await fetchCourseStructure(url);
          if (cancelled) return;
          setRootData({ type: 'course', node });
        } else if (ct === 'learning-path') {
          const node = await fetchLearningPathStructure(url);
          if (cancelled) return;
          setRootData({ type: 'learning-path', node });
        } else if (ct === 'module') {
          const node = await fetchModuleStructure(url);
          if (cancelled) return;
          setRootData({ type: 'module', node });
          setSelected(new Set(node.units.map((u) => u.url)));
        } else if (ct === 'unit') {
          const modUrl = moduleUrlFromUnitUrl(url);
          const modNode = await fetchModuleStructure(modUrl);
          if (cancelled) return;
          setRootData({ type: 'unit', unitUrl: url, moduleNode: modNode });
          setSelected(new Set(modNode.units.map((u) => u.url)));
        } else {
          setRootData(null);
        }
      } catch (e) {
        if (!cancelled)
          setErrorMsg('Failed to load page structure: ' + ((e as any)?.message ?? String(e)));
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();

    return () => { cancelled = true; };
  }, []);

  // ── Keep rootDataRef in sync ────────────────────────────────────────────────

  useEffect(() => {
    rootDataRef.current = rootData;
  }, [rootData]);

  // ── Persist state whenever the tree/selection/expansion changes ──────────────

  useEffect(() => {
    if (!currentUrl || !rootData) return;
    saveState(currentUrl, rootData, selected, expanded);
  }, [currentUrl, rootData, selected, expanded]);

  // ── Subscribe to live job-progress updates from the offscreen worker ─────────

  useEffect(() => {
    const listener = (
      changes: { [key: string]: chrome.storage.StorageChange },
      area: string
    ) => {
      if (area !== 'session' || !changes.mlfJob) return;
      const next = changes.mlfJob.newValue as JobState | undefined;
      setJobState(next ?? null);
    };
    chrome.storage.onChanged.addListener(listener);
    return () => chrome.storage.onChanged.removeListener(listener);
  }, []);

  // ── Lazy-load: LP inside a Course ───────────────────────────────────────────

  const expandLP = useCallback(async (lpUrl: string) => {
    // Skip if already loaded
    const cur = rootDataRef.current;
    if (cur?.type === 'course') {
      const lp = cur.node.learningPaths.find((l) => l.url === lpUrl);
      if (lp?.modulesLoaded) return;
    }
    setLoadingIds((prev) => new Set(prev).add(lpUrl));
    try {
      const lpData = await fetchLearningPathStructure(lpUrl);
      setRootData((prev) => {
        if (!prev || prev.type !== 'course') return prev;
        return {
          ...prev,
          node: setLPInCourse(prev.node, lpUrl, {
            title: lpData.title,
            modules: lpData.modules,
            modulesLoaded: true,
          }),
        };
      });
    } catch {
      /* silently ignore */
    } finally {
      setLoadingIds((prev) => {
        const n = new Set(prev);
        n.delete(lpUrl);
        return n;
      });
    }
  }, []);

  // ── Lazy-load: Module units ─────────────────────────────────────────────────

  const expandModule = useCallback(async (modUrl: string, parentLpUrl?: string) => {
    // Skip if already loaded
    const cur = rootDataRef.current;
    if (cur) {
      let mod: ModuleNode | undefined;
      if (cur.type === 'learning-path') {
        mod = cur.node.modules.find((m) => m.url === modUrl);
      } else if (cur.type === 'course' && parentLpUrl) {
        const lp = cur.node.learningPaths.find((l) => l.url === parentLpUrl);
        mod = lp?.modules.find((m) => m.url === modUrl);
      }
      if (mod?.unitsLoaded) return;
    }
    setLoadingIds((prev) => new Set(prev).add(modUrl));
    try {
      const modData = await fetchModuleStructure(modUrl);
      setRootData((prev) => {
        if (!prev) return prev;
        if (prev.type === 'learning-path') {
          return {
            ...prev,
            node: setModuleInLP(prev.node, modUrl, {
              units: modData.units,
              title: modData.title,
              unitsLoaded: true,
            }),
          };
        }
        if (prev.type === 'course' && parentLpUrl) {
          return {
            ...prev,
            node: setModuleInCourse(prev.node, parentLpUrl, modUrl, {
              units: modData.units,
              title: modData.title,
              unitsLoaded: true,
            }),
          };
        }
        return prev;
      });
    } catch {
      /* silently ignore */
    } finally {
      setLoadingIds((prev) => {
        const n = new Set(prev);
        n.delete(modUrl);
        return n;
      });
    }
  }, []);

  // ── Toggle expand/collapse ──────────────────────────────────────────────────

  const toggleExpanded = useCallback(
    (url: string, nodeType: 'lp' | 'module', parentLpUrl?: string) => {
      setExpanded((prev) => {
        const next = new Set(prev);
        if (next.has(url)) {
          next.delete(url);
          return next;
        }
        next.add(url);
        // Kick off lazy load when expanding
        if (nodeType === 'lp') expandLP(url);
        else expandModule(url, parentLpUrl);
        return next;
      });
    },
    [expandLP, expandModule]
  );

  // ── Selection helpers ───────────────────────────────────────────────────────

  const toggleUnit = useCallback((url: string) => {
    setSelected((prev) => {
      const next = new Set(prev);
      next.has(url) ? next.delete(url) : next.add(url);
      return next;
    });
  }, []);

  const toggleModule = useCallback(
    (mod: ModuleNode, parentLpUrl?: string) => {
      if (!mod.unitsLoaded) {
        // Load units then select all
        expandModule(mod.url, parentLpUrl).then(() => {
          setRootData((prev) => {
            let updated: ModuleNode | undefined;
            if (prev?.type === 'learning-path')
              updated = prev.node.modules.find((m) => m.url === mod.url);
            else if (prev?.type === 'course' && parentLpUrl) {
              const lp = prev.node.learningPaths.find((l) => l.url === parentLpUrl);
              updated = lp?.modules.find((m) => m.url === mod.url);
            }
            const urls = updated?.units.map((u) => u.url) ?? [];
            setSelected((s) => {
              const n = new Set(s);
              urls.forEach((u) => n.add(u));
              return n;
            });
            return prev;
          });
        });
        return;
      }
      const urls = unitsOfModule(mod);
      const cs = checkStateFor(urls, selected);
      setSelected((prev) => {
        const next = new Set(prev);
        cs === 'all' ? urls.forEach((u) => next.delete(u)) : urls.forEach((u) => next.add(u));
        return next;
      });
    },
    [selected, expandModule]
  );

  const toggleLP = useCallback(
    (lp: LearningPathNode) => {
      if (!lp.modulesLoaded) {
        expandLP(lp.url).then(() => {
          setRootData((prev) => {
            const updated =
              prev?.type === 'course'
                ? prev.node.learningPaths.find((l) => l.url === lp.url)
                : null;
            const urls = updated ? unitsOfLP(updated) : [];
            setSelected((s) => {
              const n = new Set(s);
              urls.forEach((u) => n.add(u));
              return n;
            });
            return prev;
          });
        });
        return;
      }
      const urls = unitsOfLP(lp);
      const cs = checkStateFor(urls, selected);
      setSelected((prev) => {
        const next = new Set(prev);
        cs === 'all' ? urls.forEach((u) => next.delete(u)) : urls.forEach((u) => next.add(u));
        return next;
      });
    },
    [selected, expandLP]
  );

  const selectAll = useCallback(() => {
    setSelected(new Set(allUnitsOfRoot(rootData)));
  }, [rootData]);

  const deselectAll = useCallback(() => setSelected(new Set()), []);

  // ── Resolve selected units for PDF (with lazy fetching) ────────────────────

  const resolveSelectedUnits = useCallback(async (): Promise<{ title: string; url: string }[]> => {
    if (!rootData) return [];

    const getModUnits = async (
      mod: ModuleNode,
      lpUrl?: string
    ): Promise<{ title: string; url: string }[]> => {
      let units = mod.units;
      if (!mod.unitsLoaded) {
        const fetched = await fetchModuleStructure(mod.url);
        units = fetched.units;
        setRootData((prev) => {
          if (!prev) return prev;
          if (prev.type === 'learning-path')
            return {
              ...prev,
              node: setModuleInLP(prev.node, mod.url, { units: fetched.units, unitsLoaded: true }),
            };
          if (prev.type === 'course' && lpUrl)
            return {
              ...prev,
              node: setModuleInCourse(prev.node, lpUrl, mod.url, {
                units: fetched.units,
                unitsLoaded: true,
              }),
            };
          return prev;
        });
      }
      return units.filter((u) => selected.has(u.url));
    };

    if (rootData.type === 'module')
      return rootData.node.units.filter((u) => selected.has(u.url));

    if (rootData.type === 'unit')
      return (rootData.moduleNode?.units ?? []).filter((u) => selected.has(u.url));

    if (rootData.type === 'learning-path') {
      const results = await Promise.all(
        rootData.node.modules.map((mod) => getModUnits(mod)),
      );
      return results.flat();
    }

    if (rootData.type === 'course') {
      // Parallelize LP and module resolution to avoid sequential blocking
      const lpResults = await Promise.all(
        rootData.node.learningPaths.map(async (lp) => {
          let mods = lp.modules;
          if (!lp.modulesLoaded) {
            const fetched = await fetchLearningPathStructure(lp.url);
            mods = fetched.modules;
            setRootData((prev) => {
              if (!prev || prev.type !== 'course') return prev;
              return {
                ...prev,
                node: setLPInCourse(prev.node, lp.url, {
                  title: fetched.title,
                  modules: mods,
                  modulesLoaded: true,
                }),
              };
            });
          }
          const modResults = await Promise.all(
            mods.map((mod) => getModUnits(mod, lp.url)),
          );
          return modResults.flat();
        }),
      );
      return lpResults.flat();
    }
    return [];
  }, [rootData, selected]);

  // ── Generate PDF (dispatches to background offscreen worker) ───────────────

  const handleGenerate = useCallback(async () => {
    setErrorMsg(null);
    setLoading(true); // brief resolveSelectedUnits phase
    try {
      // Race the unit-resolution against a hard timeout so we never hang at "Preparing…"
      const units = await Promise.race([
        resolveSelectedUnits(),
        new Promise<never>((_, reject) =>
          setTimeout(
            () => reject(new Error('Timed out resolving units. Please try again.')),
            30_000,
          ),
        ),
      ]);
      if (units.length === 0) {
        setErrorMsg('Please select at least one unit.');
        return;
      }

      const docTitle =
        rootData?.type === 'course'
          ? rootData.node.title
          : rootData?.type === 'learning-path'
          ? rootData.node.title
          : rootData?.type === 'module'
          ? rootData.node.title
          : rootData?.type === 'unit'
          ? rootData.moduleNode?.title ?? 'learnflow'
          : 'learnflow';

      const total = units.length;
      const pendingJob: PendingJob = { units, docTitle, filename: `${docTitle}.pdf` };
      const initialJob: JobState = {
        status: 'fetching',
        pct: 0,
        label: `Fetching 0 / ${total} units\u2026`,
        docTitle,
        totalUnits: total,
      };

      // Write job spec + initial status so both the offscreen worker and popup
      // can read them immediately.
      await chrome.storage.session.set({ mlfPendingJob: pendingJob, mlfJob: initialJob });
      setJobState(initialJob);

      // Ask the background service worker to open the offscreen document.
      const resp = await chrome.runtime.sendMessage({ type: 'MLF_START_JOB', target: 'background' });
      if (resp && !resp.ok) {
        throw new Error(resp.error ?? 'Failed to start background worker');
      }
    } catch (err) {
      setErrorMsg('Failed to start job: ' + ((err as any)?.message ?? String(err)));
    } finally {
      setLoading(false);
    }
  }, [resolveSelectedUnits, rootData]);

  // ── Cancel in-progress job ──────────────────────────────────────────────────

  const handleCancel = useCallback(() => {
    chrome.runtime.sendMessage({ type: 'MLF_CANCEL_JOB', target: 'background' }).catch(() => {});
    setJobState(null);
  }, []);

  // ── Dismiss completed / errored job banner ──────────────────────────────────

  const handleDismissJob = useCallback(() => {
    chrome.storage.session.remove('mlfJob').catch(() => {});
    setJobState(null);
  }, []);

  // ── Render helpers ──────────────────────────────────────────────────────────

  const renderUnit = (unit: UnitNode, key: string) => (
    <label key={key} className="tree-unit">
      <input
        className="checkbox"
        type="checkbox"
        checked={selected.has(unit.url)}
        onChange={() => toggleUnit(unit.url)}
      />
      <span className="tree-unit__title">{unit.title}</span>
    </label>
  );

  const renderModule = (
    mod: ModuleNode,
    isExp: boolean,
    parentLpUrl: string | undefined,
    key: string
  ) => {
    const cs = checkStateFor(unitsOfModule(mod), selected);
    const isLoadingNode = loadingIds.has(mod.url);
    return (
      <div key={key} className="tree-module">
        <div className="tree-node__header">
          <button
            className="tree-expand-btn"
            onClick={() => toggleExpanded(mod.url, 'module', parentLpUrl)}
            title={isExp ? 'Collapse' : 'Expand'}
          >
            {isLoadingNode ? (
              <Loader2 size={13} className="spin" />
            ) : isExp ? (
              <ChevronDown size={13} />
            ) : (
              <ChevronRight size={13} />
            )}
          </button>
          <TriCheckbox
            state={mod.unitsLoaded ? cs : 'none'}
            onChange={() => toggleModule(mod, parentLpUrl)}
          />
          <BadgeTag type="module" />
          <span className="tree-node__label">{mod.title}</span>
          {!mod.unitsLoaded && !isLoadingNode && (
            <span className="tree-node__hint">expand to load</span>
          )}
        </div>
        {isExp && (
          <div className="tree-node__children">
            {isLoadingNode && <div className="tree-loading">Loading units…</div>}
            {!isLoadingNode && mod.unitsLoaded && mod.units.length === 0 && (
              <div className="tree-empty">No units found.</div>
            )}
            {!isLoadingNode && mod.unitsLoaded && mod.units.map((u) => renderUnit(u, u.url))}
          </div>
        )}
      </div>
    );
  };

  const renderLP = (lp: LearningPathNode, isExp: boolean, key: string) => {
    const cs = checkStateFor(unitsOfLP(lp), selected);
    const isLoadingNode = loadingIds.has(lp.url);
    return (
      <div key={key} className="tree-lp">
        <div className="tree-node__header">
          <button
            className="tree-expand-btn"
            onClick={() => toggleExpanded(lp.url, 'lp')}
            title={isExp ? 'Collapse' : 'Expand'}
          >
            {isLoadingNode ? (
              <Loader2 size={13} className="spin" />
            ) : isExp ? (
              <ChevronDown size={13} />
            ) : (
              <ChevronRight size={13} />
            )}
          </button>
          <TriCheckbox
            state={lp.modulesLoaded ? cs : 'none'}
            onChange={() => toggleLP(lp)}
          />
          <BadgeTag type="learning-path" />
          <span className="tree-node__label">{lp.title}</span>
          {!lp.modulesLoaded && !isLoadingNode && (
            <span className="tree-node__hint">expand to load</span>
          )}
        </div>
        {isExp && (
          <div className="tree-node__children">
            {isLoadingNode && <div className="tree-loading">Loading modules…</div>}
            {!isLoadingNode && lp.modulesLoaded && lp.modules.length === 0 && (
              <div className="tree-empty">No modules found.</div>
            )}
            {!isLoadingNode &&
              lp.modules.map((mod) =>
                renderModule(mod, expanded.has(mod.url), lp.url, mod.url)
              )}
          </div>
        )}
      </div>
    );
  };

  // ── Context header ──────────────────────────────────────────────────────────

  const renderContextHeader = () => {
    if (!rootData) return null;
    let badge: ContentType = 'unknown';
    let title = '';
    let sub = '';

    if (rootData.type === 'course') {
      badge = 'course';
      title = rootData.node.title;
      sub = `${rootData.node.learningPaths.length} learning path${rootData.node.learningPaths.length !== 1 ? 's' : ''}`;
    } else if (rootData.type === 'learning-path') {
      badge = 'learning-path';
      title = rootData.node.title;
      sub = `${rootData.node.modules.length} module${rootData.node.modules.length !== 1 ? 's' : ''}`;
    } else if (rootData.type === 'module') {
      badge = 'module';
      title = rootData.node.title;
      sub = `${rootData.node.units.length} unit${rootData.node.units.length !== 1 ? 's' : ''}`;
    } else if (rootData.type === 'unit') {
      badge = 'unit';
      title = rootData.moduleNode?.title ?? 'Module';
      sub = rootData.moduleNode
        ? `${rootData.moduleNode.units.length} units — showing parent module`
        : 'Showing parent module';
    }

    return (
      <div className="context-header">
        <div className="context-header__row">
          <BadgeTag type={badge} />
          <span className="context-header__title">{title}</span>
        </div>
        {sub && <div className="context-header__sub">{sub}</div>}
      </div>
    );
  };

  // ── Tree body ───────────────────────────────────────────────────────────────

  const renderTree = () => {
    if (loading && !rootData) {
      return <div className="tree-loading tree-loading--center">Loading content structure…</div>;
    }
    if (!rootData && !loading) {
      return (
        <div className="tree-empty tree-empty--center">
          Open a Microsoft Learn course, learning path, module, or unit page.
        </div>
      );
    }
    if (!rootData) return null;

    if (rootData.type === 'module') {
      return <div className="tree">{rootData.node.units.map((u) => renderUnit(u, u.url))}</div>;
    }
    if (rootData.type === 'unit') {
      const mod = rootData.moduleNode;
      if (!mod) return <div className="tree-empty">No module data available.</div>;
      return <div className="tree">{mod.units.map((u) => renderUnit(u, u.url))}</div>;
    }
    if (rootData.type === 'learning-path') {
      return (
        <div className="tree">
          {rootData.node.modules.map((mod) =>
            renderModule(mod, expanded.has(mod.url), undefined, mod.url)
          )}
        </div>
      );
    }
    if (rootData.type === 'course') {
      return (
        <div className="tree">
          {rootData.node.learningPaths.map((lp) =>
            renderLP(lp, expanded.has(lp.url), lp.url)
          )}
        </div>
      );
    }
    return null;
  };

  // ── Stats ───────────────────────────────────────────────────────────────────

  const loadedUnitUrls = allUnitsOfRoot(rootData);
  // Only count selections that actually exist in the current tree
  const selCount = loadedUnitUrls.filter((u) => selected.has(u)).length;
  const jobRunning = jobState?.status === 'fetching' || jobState?.status === 'rendering';

  // Build contextual selection summary (LP / module / unit counts)
  const selectionSummary = (() => {
    if (!rootData) return '';
    const parts: string[] = [];

    if (rootData.type === 'course') {
      // Count distinct LPs with at least one selected unit
      const selLPs = rootData.node.learningPaths.filter((lp) =>
        unitsOfLP(lp).some((u) => selected.has(u))
      ).length;
      parts.push(`${selLPs} learning path${selLPs !== 1 ? 's' : ''}`);

      // Count distinct modules with at least one selected unit
      const selMods = rootData.node.learningPaths.flatMap((lp) => lp.modules).filter((m) =>
        unitsOfModule(m).some((u) => selected.has(u))
      ).length;
      parts.push(`${selMods} module${selMods !== 1 ? 's' : ''}`);
    }

    if (rootData.type === 'learning-path') {
      const selMods = rootData.node.modules.filter((m) =>
        unitsOfModule(m).some((u) => selected.has(u))
      ).length;
      parts.push(`${selMods} module${selMods !== 1 ? 's' : ''}`);
    }

    parts.push(`${selCount} unit${selCount !== 1 ? 's' : ''}`);
    return parts.join(' · ');
  })();

  // ── Full render ─────────────────────────────────────────────────────────────

  return (
    <div className="popup">
      <h2 className="header">MS LearnFlow</h2>
      <div className="panel">
        {renderContextHeader()}

        {rootData && (
          <div className="selection-toolbar">
            <span className="selection-toolbar__count">
              {selectionSummary} selected
            </span>
            <div className="selection-toolbar__actions">
              <button className="btn-link" onClick={selectAll} disabled={loading || jobRunning}>
                All
              </button>
              <span className="selection-toolbar__sep">·</span>
              <button className="btn-link" onClick={deselectAll} disabled={loading || jobRunning}>
                None
              </button>
            </div>
          </div>
        )}

        <div className="tree-container">{renderTree()}</div>

        {rootData && (
          <div className="controls">
            <button
              onClick={handleGenerate}
              disabled={loading || jobRunning || selCount === 0}
              className="btn btn-primary"
            >
              {loading ? (
                <Loader2 size={15} className="spin" />
              ) : (
                <Download size={15} />
              )}
              &nbsp;Generate PDF
            </button>
            {jobRunning && (
              <button
                onClick={handleCancel}
                className="btn btn-secondary"
                style={{ marginLeft: 8 }}
              >
                Cancel
              </button>
            )}
            {loading && (
              <span style={{ marginLeft: 8, color: 'var(--muted)', fontSize: 12 }}>
                Preparing…
              </span>
            )}
          </div>
        )}

        {jobRunning && jobState && (
          <div className="progress">
            <div className="progress__label">
              <span>{jobState.label}</span>
              <span>{jobState.pct}%</span>
            </div>
            <div className="progress__bar">
              <div className="progress__inner" style={{ width: `${jobState.pct}%` }} />
            </div>
          </div>
        )}

        {jobState?.status === 'done' && (
          <div className="progress">
            <div className="progress__label">
              <span>✓ {jobState.label}</span>
              <button className="btn-link" onClick={handleDismissJob} style={{ fontSize: 11 }}>Dismiss</button>
            </div>
            <div className="progress__bar">
              <div className="progress__inner" style={{ width: '100%' }} />
            </div>
          </div>
        )}

        {jobState?.status === 'error' && (
          <div className="error-msg">
            Job failed: {jobState.error ?? jobState.label}
            <button className="btn-link" onClick={handleDismissJob} style={{ marginLeft: 8, fontSize: 11 }}>Dismiss</button>
          </div>
        )}

        {errorMsg && <div className="error-msg">{errorMsg}</div>}

        {!rootData && !loading && (
          <div className="tip">
            Tip: Navigate to any Microsoft Learn course, learning path, module, or unit page and
            click the extension icon.
          </div>
        )}
      </div>
    </div>
  );
};

export default Popup;
