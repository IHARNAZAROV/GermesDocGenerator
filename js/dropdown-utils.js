'use strict';
/**
 * dropdown-utils.js — shared open/close/keyboard utility for custom dropdowns.
 *
 * Usage:
 *   const { open, close, isOpen, handleItemKey } = bindDropdown({ ... });
 *
 * @param {object}    opts
 * @param {Element}   opts.wrapper            Root element for outside-click detection
 * @param {Element}   opts.trigger            Toggle button element
 * @param {Element}   opts.menu               Menu element (toggled via .hidden)
 * @param {string}    opts.itemSelector       CSS selector for focusable list items
 * @param {string}    [opts.selectedSelector] CSS selector for the pre-selected item;
 *                                            defaults to itemSelector + '--selected'
 * @param {string}    [opts.closingClass]     CSS class for close animation (default: 'dropdown--closing')
 * @param {Element[]} [opts.extraAriaEls]     Extra elements to mirror aria-expanded on
 * @param {Function}  [opts.onBeforeOpen]     Called immediately after menu is shown
 * @param {Function}  [opts.onAfterOpenRaf]   Called inside requestAnimationFrame after open
 * @param {Function}  [opts.getScrollParent]  Returns scroll container; scroll closes the menu
 * @param {boolean}   [opts.stopPropagation]  Whether to stopPropagation on trigger click (default: false)
 * @returns {{ open: Function, close: Function, isOpen: Function, handleItemKey: Function }}
 */
function bindDropdown({
  wrapper,
  trigger,
  menu,
  itemSelector,
  selectedSelector,
  closingClass    = 'dropdown--closing',
  extraAriaEls    = [],
  onBeforeOpen,
  onAfterOpenRaf,
  getScrollParent,
  stopPropagation = false,
}) {
  const allAriaEls = [trigger, ...extraAriaEls];

  function isOpen() { return !menu.hidden; }

  function _setExpanded(v) {
    allAriaEls.forEach(el => el?.setAttribute('aria-expanded', String(v)));
  }

  function open() {
    menu.hidden = false;
    _setExpanded(true);
    onBeforeOpen?.();
    setTimeout(() => document.addEventListener('pointerdown', _onOutside), 0);
    if (getScrollParent) {
      (getScrollParent() || window).addEventListener('scroll', _onScroll, { passive: true });
    }
    requestAnimationFrame(() => {
      onAfterOpenRaf?.();
      const sel   = menu.querySelector(selectedSelector || (itemSelector + '--selected'));
      const first = menu.querySelector(itemSelector);
      (sel || first)?.focus();
    });
  }

  function close() {
    menu.classList.add(closingClass);
    setTimeout(() => {
      menu.hidden = true;
      menu.classList.remove(closingClass);
    }, 160);
    _setExpanded(false);
    document.removeEventListener('pointerdown', _onOutside);
    if (getScrollParent) {
      (getScrollParent() || window).removeEventListener('scroll', _onScroll);
    }
    trigger?.focus();
  }

  function handleItemKey(e) {
    const items = Array.from(menu.querySelectorAll(itemSelector));
    const idx   = items.indexOf(document.activeElement);
    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        items[Math.min(idx + 1, items.length - 1)]?.focus();
        break;
      case 'ArrowUp':
        e.preventDefault();
        if (idx <= 0) { close(); return; }
        items[Math.max(idx - 1, 0)]?.focus();
        break;
      case 'Enter':
      case ' ':
        e.preventDefault();
        document.activeElement?.click();
        break;
      case 'Escape':
      case 'Tab':
        e.preventDefault();
        close();
        break;
    }
  }

  function _onOutside(e) {
    if (!wrapper.contains(e.target)) close();
  }

  function _onScroll() { if (isOpen()) close(); }

  trigger.addEventListener('click', e => {
    if (stopPropagation) e.stopPropagation();
    isOpen() ? close() : open();
  });
  trigger.addEventListener('keydown', e => {
    if (['ArrowDown', 'Enter', ' '].includes(e.key)) {
      e.preventDefault();
      if (!isOpen()) open();
    } else if (e.key === 'Escape') {
      close();
    }
  });

  return { open, close, isOpen, handleItemKey };
}
