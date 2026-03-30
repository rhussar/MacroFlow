import { useMemo } from 'react';
import spriteUrl from '../../assets/imagemso-sprite.png';
import manifest from '../features/icons/imagemso-manifest.json';
import spriteMeta from '../features/icons/imagemso-meta.json';

/**
 * Renders an ImageMSO icon from the sprite sheet using CSS background-position.
 *
 * @param {string} name - The ImageMSO identifier (e.g. "ChartColumn")
 * @param {number} size - Display size in CSS pixels (default 16)
 * @param {string} className - Optional CSS class
 */
export default function ImageMsoIcon({ name, size = 16, className = '', title }) {
  const style = useMemo(() => {
    if (!manifest[name]) return null;

    const { col, row } = manifest[name];
    const { iconSize, width, height } = spriteMeta;
    const scale = size / iconSize;

    return {
      width: size,
      height: size,
      backgroundImage: `url(${spriteUrl})`,
      backgroundPosition: `${-(col * iconSize * scale)}px ${-(row * iconSize * scale)}px`,
      backgroundSize: `${width * scale}px ${height * scale}px`,
      backgroundRepeat: 'no-repeat',
      display: 'inline-block',
      flexShrink: 0,
    };
  }, [name, size]);

  if (!style) {
    // Fallback: colored placeholder square for icons not in the sprite sheet
    return (
      <div
        className={`imagemso-icon imagemso-fallback ${className}`}
        style={{
          width: size,
          height: size,
          display: 'inline-block',
          flexShrink: 0,
          borderRadius: 2,
          background: 'var(--bg-elevated)',
          border: '1px solid var(--border-subtle)',
        }}
        title={title !== undefined ? (title || undefined) : name}
      />
    );
  }

  return (
    <div
      className={`imagemso-icon ${className}`}
      style={style}
      title={title !== undefined ? (title || undefined) : name}
    />
  );
}

/** Check if the sprite sheet is available. */
export function isSpriteReady() {
  return Boolean(manifest && spriteMeta && spriteUrl);
}

/** Get all available icon names from the manifest. */
export function getAllIconNames() {
  return manifest ? Object.keys(manifest).sort() : [];
}
