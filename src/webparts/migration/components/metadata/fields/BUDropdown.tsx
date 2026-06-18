import * as React from 'react';
import { Term } from '../../../utils/termStore';

export interface IBUDropdownProps {
  buTerms: Term[];
  selectedBUs: Term[];
  onChange: (selected: Term[]) => void;
  onClear?: () => void;
  required?: boolean;
  density?: 'default' | 'compact';
}

const containerStyle: React.CSSProperties = {
  position: 'relative'
};

const labelStyle: React.CSSProperties = {
  display: 'block',
  marginBottom: '8px',
  fontSize: '15px',
  fontWeight: 600,
  color: '#231F20'
};

const triggerStyle: React.CSSProperties = {
  width: '100%',
  minHeight: '40px',
  height: 'auto',
  padding: '8px 16px',
  border: '1px solid #8f939a',
  borderRadius: '7px',
  backgroundColor: '#ffffff',
  color: '#1d1d1f',
  textAlign: 'left',
  cursor: 'pointer',
  fontSize: '16px',
  lineHeight: 1.35,
  fontWeight: 400,
  overflow: 'visible',
  boxSizing: 'border-box',
  whiteSpace: 'normal'
};

const placeholderStyle: React.CSSProperties = {
  color: '#6b7280'
};

const menuStyle: React.CSSProperties = {
  position: 'absolute',
  top: 'calc(100% + 6px)',
  left: 0,
  right: 0,
  width: '100%',
  zIndex: 1000,
  maxHeight: '240px',
  overflowY: 'auto',
  border: '1px solid #d1d5db',
  borderRadius: '8px',
  backgroundColor: '#ffffff',
  boxShadow: '0 12px 28px rgba(15, 23, 42, 0.12)',
  padding: '8px 0'
};

const optionStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: '8px',
  padding: '8px 12px',
  fontSize: '14px',
  color: '#111827',
  cursor: 'pointer'
};

const chipWrapStyle: React.CSSProperties = {
  display: 'flex',
  flexWrap: 'wrap',
  gap: '6px',
  width: '100%',
  minWidth: 0
};

const chipStyle: React.CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  gap: '6px',
  padding: '4px 10px',
  borderRadius: '6px',
  background: '#f3f4f6',
  color: '#1d1d1f',
  fontSize: '13px',
  fontWeight: 500,
  maxWidth: '100%',
  boxSizing: 'border-box',
  lineHeight: 1.35
};

const chipRemoveStyle: React.CSSProperties = {
  border: 'none',
  background: 'transparent',
  color: '#6b7280',
  cursor: 'pointer',
  fontSize: '16px',
  lineHeight: 1,
  padding: 0,
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center'
};

export const BUDropdown: React.FC<IBUDropdownProps> = ({ buTerms, selectedBUs, onChange, onClear, required, density = 'default' }) => {
  const [isOpen, setIsOpen] = React.useState<boolean>(false);
  const containerRef = React.useRef<HTMLDivElement>(null);
  const isCompact = density === 'compact';
  const resolvedLabelStyle = React.useMemo<React.CSSProperties>(() => ({
    ...labelStyle,
    marginBottom: isCompact ? '6px' : labelStyle.marginBottom,
    fontSize: isCompact ? '11.2px' : labelStyle.fontSize,
    lineHeight: isCompact ? '16px' : labelStyle.lineHeight,
    fontWeight: 600
  }), [isCompact]);
  const resolvedTriggerStyle = React.useMemo<React.CSSProperties>(() => ({
    ...triggerStyle,
    fontSize: isCompact ? '13.6px' : triggerStyle.fontSize,
    lineHeight: isCompact ? 1.45 : triggerStyle.lineHeight
  }), [isCompact]);
  const resolvedOptionStyle = React.useMemo<React.CSSProperties>(() => ({
    ...optionStyle,
    fontSize: isCompact ? '13.6px' : optionStyle.fontSize
  }), [isCompact]);
  const resolvedChipStyle = React.useMemo<React.CSSProperties>(() => ({
    ...chipStyle,
    fontSize: isCompact ? '13.6px' : chipStyle.fontSize,
    lineHeight: isCompact ? 1.45 : chipStyle.lineHeight
  }), [isCompact]);

  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent): void => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const rootBUTerms = React.useMemo<Term[]>(
    () => buTerms.filter((term) => Array.isArray(term.children)),
    [buTerms]
  );

  const toggleSelection = (term: Term): void => {
    const isSelected = selectedBUs.some((selectedTerm) => selectedTerm.id === term.id);

    onChange(
      isSelected
        ? selectedBUs.filter((selectedTerm) => selectedTerm.id !== term.id)
        : [...selectedBUs, term]
    );
  };

  const selectedText = selectedBUs.length > 0
    ? selectedBUs.map((term) => term.name).join('; ')
    : 'Select Business Unit(s)...';

  return (
    <div ref={containerRef} style={containerStyle}>
      <label style={resolvedLabelStyle}>
        Business Unit
        {required && (
          <span style={{ color: '#d32f2f', marginLeft: '4px', fontWeight: 600 }}>*</span>
        )}
      </label>
      <div style={{ position: 'relative' }}>
        <button
          type="button"
          style={resolvedTriggerStyle}
          onClick={() => setIsOpen((currentValue) => !currentValue)}
        >
          {selectedBUs.length > 0 ? (
            <span style={chipWrapStyle}>
              {selectedBUs.map((term) => (
                <span key={term.id} style={resolvedChipStyle}>
                  <span style={{ minWidth: 0, overflowWrap: 'anywhere' }}>{term.name}</span>
                  <button
                    type="button"
                    aria-label={`Remove ${term.name}`}
                    onClick={(event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      toggleSelection(term);
                    }}
                    style={chipRemoveStyle}
                  >
                    ×
                  </button>
                </span>
              ))}
            </span>
          ) : (
            <span style={placeholderStyle}>{selectedText}</span>
          )}
        </button>
        {selectedBUs.length > 1 && onClear && (
          <button
            type="button"
            aria-label="Clear Business Unit"
            onClick={(event) => {
              event.preventDefault();
              event.stopPropagation();
              onClear();
              setIsOpen(false);
            }}
            style={{
              position: 'absolute',
              top: '10px',
              right: '10px',
              transform: 'none',
              width: '24px',
              height: '24px',
              border: 'none',
              borderRadius: '999px',
              background: 'rgba(15, 33, 51, 0.06)',
              color: '#475569',
              cursor: 'pointer',
              fontSize: '14px',
              lineHeight: 1
            }}
          >
            ×
          </button>
        )}
      </div>
      {isOpen && (
        <div style={menuStyle}>
          {rootBUTerms.map((term) => {
            const checked = selectedBUs.some((selectedTerm) => selectedTerm.id === term.id);

            return (
              <label key={term.id} style={resolvedOptionStyle}>
                <input
                  type="checkbox"
                  checked={checked}
                  onChange={() => toggleSelection(term)}
                />
                <span>{term.name}</span>
              </label>
            );
          })}
        </div>
      )}
    </div>
  );
};

export default BUDropdown;

