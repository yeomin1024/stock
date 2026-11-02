// VERSION: v1.0.0 — 2026-11-02 — Initial implementation: control panel (views, grid, layers, demo) + caption
import type { GroupName } from './building/buildBuilding';
import type { ViewName } from './stage';

export interface UIOptions {
  version: string;
  groups: readonly GroupName[];
  views: { name: ViewName; label: string }[];
  gridVisible: boolean;
  summary: string[];
  captureMode: boolean;
}

export interface UIHandlers {
  onGridToggle(visible: boolean): void;
  onLayerToggle(group: GroupName, visible: boolean): void;
  onView(view: ViewName): void;
  onDemo(): void;
}

export interface UI {
  setCaption(text: string | null): void;
  setDemoPlaying(playing: boolean): void;
  setGridChecked(checked: boolean): void;
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, props: Partial<HTMLElementTagNameMap[K]> = {}, children: (Node | string)[] = []): HTMLElementTagNameMap[K] {
  const node = Object.assign(document.createElement(tag), props);
  node.append(...children);
  return node;
}

function checkbox(label: string, checked: boolean, onChange: (v: boolean) => void, extra?: Node): HTMLLabelElement {
  const input = el('input', { type: 'checkbox', checked });
  input.addEventListener('change', () => onChange(input.checked));
  return el('label', { className: 'check' }, extra ? [input, label, extra] : [input, label]);
}

export function createUI(root: HTMLElement, opts: UIOptions, h: UIHandlers): UI {
  const caption = el('div', { className: 'caption', hidden: true });
  root.append(caption);

  if (opts.captureMode) {
    // Video capture: no controls, only a title block + caption.
    root.classList.add('capture');
    root.append(el('div', { className: 'title-block' }, [el('h1', { textContent: 'SEOUL MODULAR BUILDING' }), el('p', { textContent: opts.summary.join('  ·  ') })]));
  }

  const gridInput = el('input', { type: 'checkbox', checked: opts.gridVisible });
  gridInput.addEventListener('change', () => h.onGridToggle(gridInput.checked));

  const demoBtn = el('button', { type: 'button', className: 'primary', textContent: '▶ Assembly demo' });
  demoBtn.addEventListener('click', () => h.onDemo());

  const viewButtons = opts.views.map((v) => {
    const b = el('button', { type: 'button', textContent: v.label, title: `${v.label} view` });
    b.addEventListener('click', () => h.onView(v.name));
    return b;
  });

  const details = el('details', { open: !window.matchMedia('(max-width: 640px)').matches }, [
    el('summary', {}, [el('span', { className: 'name', textContent: 'SEOUL MODULAR BUILDING' }), el('span', { className: 'ver', textContent: `v${opts.version}` })]),
    el('ul', { className: 'summary' }, opts.summary.map((s) => el('li', { textContent: s }))),
    el('section', {}, [el('h2', { textContent: 'View' }), el('div', { className: 'row' }, viewButtons)]),
    el('section', {}, [
      el('h2', { textContent: 'Dev helpers' }),
      el('label', { className: 'check' }, [gridInput, 'GridHelper (3.0 m bay) ', el('kbd', { textContent: 'G' })]),
    ]),
    el(
      'section',
      {},
      [el('h2', { textContent: 'Layers (BuildingRoot)' }), ...opts.groups.map((g) => checkbox(g, true, (v) => h.onLayerToggle(g, v)))],
    ),
    demoBtn,
  ]);
  const panel = el('aside', { className: 'panel' }, [details]);
  if (!opts.captureMode) root.append(panel);

  window.addEventListener('keydown', (e) => {
    if ((e.key === 'g' || e.key === 'G') && !(e.target instanceof HTMLInputElement)) {
      gridInput.checked = !gridInput.checked;
      h.onGridToggle(gridInput.checked);
    }
  });

  return {
    setCaption(text) {
      caption.hidden = text === null;
      caption.textContent = text ?? '';
    },
    setDemoPlaying(playing) {
      demoBtn.textContent = playing ? '■ Stop demo' : '▶ Assembly demo';
    },
    setGridChecked(checked) {
      gridInput.checked = checked;
    },
  };
}
