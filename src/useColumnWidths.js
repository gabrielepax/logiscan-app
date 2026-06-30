import { useState, useCallback, useRef, useEffect } from 'react';

export function useColumnWidths(storageKey, defaults) {
  const [widths, setWidths] = useState(() => {
    try {
      const saved = localStorage.getItem(storageKey);
      return saved ? { ...defaults, ...JSON.parse(saved) } : { ...defaults };
    } catch { return { ...defaults }; }
  });

  const widthsRef = useRef(widths);
  useEffect(() => { widthsRef.current = widths; }, [widths]);

  const startResize = useCallback((col, e) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startWidth = widthsRef.current[col] ?? 100;
    let moved = false;

    const onMove = (moveE) => {
      if (Math.abs(moveE.clientX - startX) > 2) moved = true;
      const newWidth = Math.max(40, startWidth + moveE.clientX - startX);
      setWidths(prev => ({ ...prev, [col]: newWidth }));
    };

    const onUp = (upE) => {
      const newWidth = Math.max(40, startWidth + upE.clientX - startX);
      const updated = { ...widthsRef.current, [col]: newWidth };
      setWidths(updated);
      localStorage.setItem(storageKey, JSON.stringify(updated));
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);

      if (moved) {
        const blockClick = (ev) => {
          ev.stopPropagation();
          document.removeEventListener('click', blockClick, true);
        };
        document.addEventListener('click', blockClick, true);
      }
    };

    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }, [storageKey]);

  const resetWidths = useCallback(() => {
    setWidths({ ...defaults });
    localStorage.removeItem(storageKey);
  }, [defaults, storageKey]);

  return { widths, startResize, resetWidths };
}
