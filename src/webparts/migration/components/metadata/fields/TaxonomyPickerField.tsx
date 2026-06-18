import * as React from 'react';
import {
  buildTaxonomyTree,
  getTaxonomyPathSegments,
  ITaxonomyTerm,
  ITaxonomyTreeNode
} from '../../../services/TaxonomyService';

export interface ITaxonomyPickerFieldProps {
  className?: string;
  hierarchical?: boolean;
  hierarchicalLabels?: {
    level1: string;
    level2: string;
    level3: string;
  };
  id: string;
  multiSelect?: boolean;
  options: ITaxonomyTerm[];
  placeholder: string;
  value: ITaxonomyTerm | ITaxonomyTerm[] | null | undefined;
  onChange: (term: ITaxonomyTerm | ITaxonomyTerm[] | null) => void;
}

export const TaxonomyPickerField: React.FC<ITaxonomyPickerFieldProps> = ({
  className,
  hierarchical,
  hierarchicalLabels,
  id,
  multiSelect,
  options,
  placeholder,
  value,
  onChange
}) => {
  const containerRef = React.useRef<HTMLDivElement | null>(null);
  const [isOpen, setIsOpen] = React.useState(false);
  const [expandedKeys, setExpandedKeys] = React.useState<Record<string, boolean>>({});
  const tree = React.useMemo(() => buildTaxonomyTree(options), [options]);
  const selectedTerms = React.useMemo(
    () => (Array.isArray(value) ? value.filter(Boolean) : value ? [value] : []),
    [value]
  );
  const selectedPathSegments = React.useMemo(
    () => (!Array.isArray(value) ? getTaxonomyPathSegments(value) : []),
    [value]
  );
  const selectedKey = selectedPathSegments.join(' > ');
  const selectedIds = React.useMemo(() => {
    const ids: Record<string, boolean> = {};
    selectedTerms.forEach((term) => {
      if (term?.id) {
        ids[term.id] = true;
      }
    });
    return ids;
  }, [selectedTerms]);
  const hasValue = Array.isArray(value) ? value.length > 0 : !!value;
  const clearButtonStyle: React.CSSProperties = {
    position: 'absolute',
    top: '50%',
    right: '10px',
    transform: 'translateY(-50%)',
    width: '28px',
    height: '28px',
    border: 'none',
    borderRadius: '999px',
    background: '#eef2f7',
    color: '#64748b',
    cursor: 'pointer',
    fontSize: '16px',
    lineHeight: 1,
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center'
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

  React.useEffect(() => {
    const handleDocumentClick = (event: MouseEvent) => {
      if (!containerRef.current || containerRef.current.contains(event.target as Node)) {
        return;
      }

      setIsOpen(false);
    };

    document.addEventListener('mousedown', handleDocumentClick);
    return () => document.removeEventListener('mousedown', handleDocumentClick);
  }, []);

  React.useEffect(() => {
    if (!hierarchical || !selectedKey) {
      return;
    }

    const nextExpandedKeys: Record<string, boolean> = {};
    for (let index = 0; index < selectedPathSegments.length - 1; index++) {
      const key = selectedPathSegments.slice(0, index + 1).join(' > ');
      nextExpandedKeys[key] = true;
    }

    setExpandedKeys((currentExpandedKeys) => ({
      ...currentExpandedKeys,
      ...nextExpandedKeys
    }));
  }, [hierarchical, selectedKey, selectedPathSegments]);

  const handleFlatChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
    if (multiSelect) {
      const selectedTermsForMulti = Array.from(event.target.selectedOptions)
        .map((option) => option.value)
        .filter(Boolean)
        .map((selectedId) => options.find((option) => option.id === selectedId))
        .filter((term): term is ITaxonomyTerm => !!term);

      onChange(selectedTermsForMulti);
      return;
    }

    const selectedId = event.target.value;

    if (!selectedId) {
      onChange(null);
      return;
    }

    for (let index = 0; index < options.length; index++) {
      if (options[index].id === selectedId) {
        onChange(options[index]);
        return;
      }
    }

    onChange(null);
  };

  if (hierarchical) {
    const labels = hierarchicalLabels || {
      level1: 'Business Unit',
      level2: 'Department',
      level3: 'Sub-Department'
    };

    const selectedDisplayValue = selectedPathSegments.length > 0
      ? selectedPathSegments.join(' > ')
      : '';

    const toggleNode = (nodeKey: string) => {
      setExpandedKeys((currentExpandedKeys) => ({
        ...currentExpandedKeys,
        [nodeKey]: !currentExpandedKeys[nodeKey]
      }));
    };

    const renderNode = (node: ITaxonomyTreeNode, depth: number): React.ReactNode => {
      const hasChildren = node.children.length > 0;
      const isExpanded = !!expandedKeys[node.key];
      const isSelected = selectedKey === node.key;
      const levelLabel = depth === 0
        ? labels.level1
        : depth === 1
          ? labels.level2
          : labels.level3;

      return (
        <div key={node.key}>
          <div
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              padding: '8px 10px 8px 12px',
              paddingLeft: `${12 + depth * 18}px`,
              borderRadius: '8px',
              background: isSelected ? '#eef2ff' : 'transparent'
            }}
          >
            {hasChildren ? (
              <button
                type="button"
                onClick={() => toggleNode(node.key)}
                aria-label={`${isExpanded ? 'Collapse' : 'Expand'} ${node.label}`}
                title={`${isExpanded ? 'Collapse' : 'Expand'} ${levelLabel}`}
                style={{
                  width: '28px',
                  height: '28px',
                  borderRadius: '6px',
                  border: '1px solid rgba(15,33,51,0.1)',
                  background: '#ffffff',
                  color: '#1b3752',
                  cursor: 'pointer',
                  display: 'inline-flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  flexShrink: 0
                }}
              >
                <span
                  style={{
                    display: 'inline-block',
                    transform: isExpanded ? 'rotate(90deg)' : 'rotate(0deg)',
                    transition: 'transform 0.12s ease'
                  }}
                >
                  ▶
                </span>
              </button>
            ) : (
              <span style={{ width: '28px', flexShrink: 0 }} />
            )}

            <button
              type="button"
              onClick={() => {
                if (node.term) {
                  onChange(node.term);
                  setIsOpen(false);
                  return;
                }

                if (hasChildren) {
                  toggleNode(node.key);
                }
              }}
              style={{
                flex: 1,
                textAlign: 'left',
                border: 'none',
                background: 'transparent',
                color: '#0f2133',
                fontWeight: isSelected ? 700 : 600,
                fontSize: '14px',
                cursor: 'pointer',
                padding: 0
              }}
            >
              {node.label}
            </button>

            {node.term && (
              <button
                type="button"
                onClick={() => {
                  onChange(node.term);
                  setIsOpen(false);
                }}
                style={{
                  border: '1px solid rgba(15,33,51,0.1)',
                  background: isSelected ? '#dbe4ff' : '#ffffff',
                  color: '#1b3752',
                  borderRadius: '999px',
                  padding: '4px 10px',
                  fontSize: '12px',
                  fontWeight: 700,
                  cursor: 'pointer',
                  flexShrink: 0
                }}
              >
                Select
              </button>
            )}
          </div>

          {hasChildren && isExpanded && (
            <div>{node.children.map((childNode) => renderNode(childNode, depth + 1))}</div>
          )}
        </div>
      );
    };

    return (
      <div ref={containerRef} style={{ position: 'relative' }}>
        <div style={{ position: 'relative' }}>
          <button
            id={id}
            type="button"
            className={className}
            onClick={() => setIsOpen((currentIsOpen) => !currentIsOpen)}
            style={{
              textAlign: 'left',
              width: '100%',
              paddingRight: hasValue ? '48px' : undefined,
              backgroundImage: hasValue ? 'none' : undefined
            }}
          >
            {selectedDisplayValue || placeholder}
          </button>
          {hasValue && (
            <button
              type="button"
              aria-label={`Clear ${placeholder}`}
              onClick={(event) => {
                event.preventDefault();
                event.stopPropagation();
                onChange(null);
                setIsOpen(false);
              }}
              style={clearButtonStyle}
            >
              ×
            </button>
          )}
        </div>

        {isOpen && (
          <div
            style={{
              position: 'absolute',
              top: 'calc(100% + 6px)',
              left: 0,
              right: 0,
              width: '100%',
              background: '#fff',
              border: '1px solid rgba(15,33,51,0.12)',
              borderRadius: '10px',
              boxShadow: '0 12px 30px rgba(15,33,51,0.12)',
              zIndex: 1000,
              padding: '10px'
            }}
          >
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                gap: '12px',
                marginBottom: '10px'
              }}
            >
              <div style={{ fontSize: '12px', fontWeight: 700, color: '#5b6b7a' }}>
                {labels.level1} / {labels.level2} / {labels.level3}
              </div>

              {value && (
                <button
                  type="button"
                  onClick={() => onChange(null)}
                  style={{
                    border: 'none',
                    background: 'transparent',
                    color: '#6b7280',
                    fontSize: '12px',
                    fontWeight: 700,
                    cursor: 'pointer'
                  }}
                >
                  Clear
                </button>
              )}
            </div>

            <div
              role="tree"
              aria-label={placeholder}
              style={{
                maxHeight: '280px',
                overflowY: 'auto',
                border: '1px solid rgba(15,33,51,0.08)',
                borderRadius: '8px',
                padding: '6px'
              }}
            >
              {tree.length > 0 ? tree.map((node) => renderNode(node, 0)) : (
                <div style={{ padding: '12px', color: '#6b7280', fontSize: '13px' }}>
                  No terms available.
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    );
  }

  if (multiSelect) {
    const selectedDisplayValue = selectedTerms.length > 0
      ? selectedTerms.map((term) => term.label).join(', ')
      : '';

    const toggleTerm = (term: ITaxonomyTerm) => {
      const isSelected = !!selectedIds[term.id];
      const nextTerms = isSelected
        ? selectedTerms.filter((selectedTerm) => selectedTerm.id !== term.id)
        : selectedTerms.concat(term);

      onChange(nextTerms);
    };

    return (
      <div ref={containerRef} style={{ position: 'relative' }}>
        <div style={{ position: 'relative' }}>
          <button
            id={id}
            type="button"
            className={className}
            onClick={() => setIsOpen((currentIsOpen) => !currentIsOpen)}
            style={{
              textAlign: 'left',
              width: '100%',
              height: 'auto',
              minHeight: '40px',
              overflow: 'visible',
              whiteSpace: 'normal',
              color: selectedTerms.length > 0 ? '#1d1d1f' : undefined,
              paddingRight: selectedTerms.length > 1 ? '48px' : hasValue ? '40px' : undefined,
              backgroundImage: hasValue ? 'none' : undefined
            }}
          >
            {selectedTerms.length > 0 ? (
              <span style={chipWrapStyle}>
                {selectedTerms.map((term) => (
                  <span key={term.id} style={chipStyle}>
                    <span style={{ minWidth: 0, overflowWrap: 'anywhere' }}>{term.label}</span>
                    <button
                      type="button"
                      aria-label={`Remove ${term.label}`}
                      onClick={(event) => {
                        event.preventDefault();
                        event.stopPropagation();
                        toggleTerm(term);
                      }}
                      style={chipRemoveStyle}
                    >
                      ×
                    </button>
                  </span>
                ))}
              </span>
            ) : (
              selectedDisplayValue || placeholder
            )}
          </button>
          {selectedTerms.length > 1 && (
            <button
              type="button"
              aria-label={`Clear ${placeholder}`}
              onClick={(event) => {
                event.preventDefault();
                event.stopPropagation();
                onChange([]);
                setIsOpen(false);
              }}
              style={clearButtonStyle}
            >
              ×
            </button>
          )}
        </div>

        {isOpen && (
          <div
            style={{
              position: 'absolute',
              top: 'calc(100% + 6px)',
              left: 0,
              right: 0,
              width: '100%',
              background: '#fff',
              border: '1px solid rgba(15,33,51,0.12)',
              borderRadius: '10px',
              boxShadow: '0 12px 30px rgba(15,33,51,0.12)',
              zIndex: 1000,
              padding: '10px'
            }}
          >
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                gap: '12px',
                marginBottom: '10px'
              }}
            >
              <div style={{ fontSize: '12px', fontWeight: 700, color: '#5b6b7a' }}>
                Select one or more
              </div>

              {selectedTerms.length > 0 && (
                <button
                  type="button"
                  onClick={() => onChange([])}
                  style={{
                    border: 'none',
                    background: 'transparent',
                    color: '#6b7280',
                    fontSize: '12px',
                    fontWeight: 700,
                    cursor: 'pointer'
                  }}
                >
                  Clear
                </button>
              )}
            </div>

            <div
              role="listbox"
              aria-multiselectable="true"
              aria-label={placeholder}
              style={{
                maxHeight: '280px',
                overflowY: 'auto',
                border: '1px solid rgba(15,33,51,0.08)',
                borderRadius: '8px',
                padding: '6px'
              }}
            >
              {options.length > 0 ? options.map((option) => {
                const isSelected = !!selectedIds[option.id];

                return (
                  <label
                    key={option.id}
                    style={{
                      display: 'flex',
                      alignItems: 'flex-start',
                      gap: '10px',
                      padding: '8px 10px',
                      borderRadius: '8px',
                      cursor: 'pointer',
                      background: isSelected ? '#eef2ff' : 'transparent'
                    }}
                  >
                    <input
                      type="checkbox"
                      checked={isSelected}
                      onChange={() => toggleTerm(option)}
                      style={{ marginTop: '2px' }}
                    />
                    <span style={{ color: '#0f2133', fontSize: '14px', fontWeight: isSelected ? 700 : 500 }}>
                      {option.path || option.label}
                    </span>
                  </label>
                );
              }) : (
                <div style={{ padding: '12px', color: '#6b7280', fontSize: '13px' }}>
                  No terms available.
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    );
  }

  const singleValue = Array.isArray(value) ? null : value;

  return (
    <select id={id} value={singleValue?.id || ''} onChange={handleFlatChange} className={className}>
      <option value="">{placeholder}</option>
      {options.map((option) => (
        <option key={option.id} value={option.id}>
          {option.path || option.label}
        </option>
      ))}
    </select>
  );
};

export default TaxonomyPickerField;

