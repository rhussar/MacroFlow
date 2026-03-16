import React, { useState, useCallback, useEffect, useRef } from 'react';

const SplitDivider = ({ onResize }) => {
  const [dragging, setDragging] = useState(false);
  const dividerRef = useRef(null);

  const onMouseDown = useCallback((e) => {
    e.preventDefault();
    setDragging(true);
  }, []);

  useEffect(() => {
    if (!dragging) return;

    const onMouseMove = (e) => {
      const container = dividerRef.current?.parentElement;
      if (!container) return;
      const rect = container.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const pct = Math.min(Math.max((x / rect.width) * 100, 15), 85);
      onResize(pct);
    };

    const onMouseUp = () => setDragging(false);

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';

    return () => {
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    };
  }, [dragging, onResize]);

  return (
    <div
      ref={dividerRef}
      className={`split-divider ${dragging ? 'dragging' : ''}`}
      onMouseDown={onMouseDown}
    />
  );
};

export default SplitDivider;
