import * as React from 'react';
import { Term } from '../../../utils/termStore';

export interface IDeptSubDeptDropdownProps {
  buTerms?: Term[];
  selectedBUs: Term[];
  selectedDepts: Term[];
  onChange: (selected: Term[]) => void;
  onClear?: () => void;
  density?: 'default' | 'compact';
}

interface IAvailableOption {
  term: Term;
  indent: 0 | 1;
  parentId: string;
  parentName: string;
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

const triggerBaseStyle: React.CSSProperties = {
  width: '100%',
  minHeight: '40px',
  height: 'auto',
  padding: '8px 16px',
  border: '1px solid #8f939a',
  borderRadius: '7px',
  backgroundColor: '#ffffff',
  color: '#1d1d1f',
  textAlign: 'left',
  fontSize: '16px',
  lineHeight: 1.35,
  fontWeight: 400,
  overflow: 'visible',
  boxSizing: 'border-box',
  whiteSpace: 'normal'
};

const enabledTriggerStyle: React.CSSProperties = {
  ...triggerBaseStyle,
  cursor: 'pointer'
};

const disabledTriggerStyle: React.CSSProperties = {
  ...triggerBaseStyle,
  backgroundColor: '#f3f4f6',
  color: '#9ca3af',
  cursor: 'not-allowed'
};

const placeholderStyle: React.CSSProperties = {
  color: '#6b7280'
};

const emptyStateStyle: React.CSSProperties = {
  padding: '10px 12px',
  fontSize: '13px',
  color: '#6b7280'
};

const menuStyle: React.CSSProperties = {
  position: 'absolute',
  top: 'calc(100% + 6px)',
  left: 0,
  right: 0,
  width: '100%',
  zIndex: 1000,
  maxHeight: '280px',
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

const dropdownCaretStyle: React.CSSProperties = {
  position: 'absolute',
  top: '50%',
  right: '14px',
  transform: 'translateY(-50%)',
  width: 0,
  height: 0,
  borderLeft: '6px solid transparent',
  borderRight: '6px solid transparent',
  borderTop: '7px solid #64748b',
  pointerEvents: 'none'
};

export const DeptSubDeptDropdown: React.FC<IDeptSubDeptDropdownProps> = ({
  buTerms,
  selectedBUs,
  selectedDepts,
  onChange,
  onClear,
  density = 'default'
}) => {
  const [isOpen, setIsOpen] = React.useState<boolean>(false);
  const containerRef = React.useRef<HTMLDivElement>(null);
  const isCompact = density === 'compact';
  const hasSelectedBUs = selectedBUs.length > 0;
  const resolvedLabelStyle = React.useMemo<React.CSSProperties>(() => ({
    ...labelStyle,
    marginBottom: isCompact ? '6px' : labelStyle.marginBottom,
    fontSize: isCompact ? '11.2px' : labelStyle.fontSize,
    lineHeight: isCompact ? '16px' : labelStyle.lineHeight,
    fontWeight: 600
  }), [isCompact]);
  const resolvedTriggerStyle = React.useMemo<React.CSSProperties>(() => ({
    ...(hasSelectedBUs ? enabledTriggerStyle : disabledTriggerStyle),
    fontSize: isCompact ? '13.6px' : triggerBaseStyle.fontSize,
    lineHeight: isCompact ? 1.45 : triggerBaseStyle.lineHeight,
    paddingRight: selectedDepts.length > 1 ? '72px' : '44px'
  }), [hasSelectedBUs, isCompact, selectedDepts.length]);
  const resolvedOptionStyle = React.useMemo<React.CSSProperties>(() => ({
    ...optionStyle,
    fontSize: isCompact ? '13.6px' : optionStyle.fontSize
  }), [isCompact]);
  const resolvedEmptyStateStyle = React.useMemo<React.CSSProperties>(() => ({
    ...emptyStateStyle,
    fontSize: isCompact ? '13.6px' : emptyStateStyle.fontSize
  }), [isCompact]);
  const resolvedChipStyle = React.useMemo<React.CSSProperties>(() => ({
    ...chipStyle,
    fontSize: isCompact ? '13.6px' : chipStyle.fontSize,
    lineHeight: isCompact ? 1.45 : chipStyle.lineHeight
  }), [isCompact]);

  const resolvedSelectedBUs = React.useMemo<Term[]>(() => {
    const sourceTerms = Array.isArray(buTerms) ? buTerms : [];
    const sourceMap = sourceTerms.reduce<Record<string, Term>>((accumulator, term) => {
      accumulator[term.id] = term;
      return accumulator;
    }, {});

    return selectedBUs
      .map((selectedBU) => sourceMap[selectedBU.id] || selectedBU)
      .filter(Boolean);
  }, [buTerms, selectedBUs]);

  const availableOptions = React.useMemo<IAvailableOption[]>(() => {
    const options: IAvailableOption[] = [];

    resolvedSelectedBUs.forEach((bu) => {
      (bu.children || []).forEach((dept) => {
        options.push({
          term: dept,
          indent: 0,
          parentId: dept.id,
          parentName: dept.name
        });

        (dept.children || []).forEach((subDept) => {
          options.push({
            term: subDept,
            indent: 1,
            parentId: dept.id,
            parentName: dept.name
          });
        });
      });
    });

    return options;
  }, [resolvedSelectedBUs]);

  const availableOptionMap = React.useMemo<Record<string, IAvailableOption>>(
    () => availableOptions.reduce<Record<string, IAvailableOption>>((accumulator, option) => {
      accumulator[option.term.id] = option;
      return accumulator;
    }, {}),
    [availableOptions]
  );

  const subDeptIdsByParentId = React.useMemo<Record<string, string[]>>(
    () => availableOptions.reduce<Record<string, string[]>>((accumulator, option) => {
      if (option.indent === 1) {
        if (!accumulator[option.parentId]) {
          accumulator[option.parentId] = [];
        }

        accumulator[option.parentId].push(option.term.id);
      }

      return accumulator;
    }, {}),
    [availableOptions]
  );

  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent): void => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  React.useEffect(() => {
    const allowedIds = new Set(availableOptions.map((option) => option.term.id));
    const filteredSelection = selectedDepts.filter((term) => allowedIds.has(term.id));

    if (filteredSelection.length !== selectedDepts.length) {
      onChange(filteredSelection);
    }

    if (selectedBUs.length === 0) {
      setIsOpen(false);
    }
  }, [availableOptions, onChange, selectedBUs.length, selectedDepts]);

  const toggleSelection = (option: IAvailableOption): void => {
    const isSelected = selectedDepts.some((selectedTerm) => selectedTerm.id === option.term.id);

    if (!isSelected) {
      if (option.indent === 1) {
        const parentSelected = selectedDepts.some((selectedTerm) => selectedTerm.id === option.parentId);
        const nextSelection = parentSelected
          ? [...selectedDepts, option.term]
          : [availableOptionMap[option.parentId].term, ...selectedDepts, option.term];

        onChange(nextSelection);
        return;
      }

      onChange([...selectedDepts, option.term]);
      return;
    }

    if (option.indent === 0) {
      const childIds = new Set(subDeptIdsByParentId[option.term.id] || []);
      onChange(
        selectedDepts.filter((selectedTerm) => selectedTerm.id !== option.term.id && !childIds.has(selectedTerm.id))
      );
      return;
    }

    onChange(selectedDepts.filter((selectedTerm) => selectedTerm.id !== option.term.id));
  };

  const displayText = React.useMemo(() => {
    if (selectedDepts.length === 0) {
      if (selectedBUs.length > 0 && availableOptions.length === 0) {
        return 'No departments or sub departments available for this Business Unit';
      }

      return selectedBUs.length > 0 ? 'Select Department(s)...' : 'Select a Business Unit first...';
    }

    const grouped = selectedDepts.reduce<Record<string, { deptName: string; hasDept: boolean; subs: string[] }>>(
      (accumulator, term) => {
        const option = availableOptionMap[term.id];
        if (!option) {
          return accumulator;
        }

        if (!accumulator[option.parentId]) {
          accumulator[option.parentId] = {
            deptName: option.parentName,
            hasDept: false,
            subs: []
          };
        }

        if (option.indent === 0) {
          accumulator[option.parentId].hasDept = true;
        } else {
          accumulator[option.parentId].subs.push(option.term.name);
        }

        return accumulator;
      },
      {}
    );

    return Object.keys(grouped)
      .map((parentId) => {
        const group = grouped[parentId];
        return group.subs.length > 0
          ? `${group.deptName} > ${group.subs.join(' | ')}`
          : group.deptName;
      })
      .join('; ');
  }, [availableOptionMap, availableOptions.length, selectedBUs.length, selectedDepts]);

  const hasAvailableOptions = availableOptions.length > 0;
  const selectedChipItems = React.useMemo(() => {
    const groupedChips = new Map<string, {
      id: string;
      deptName: string;
      subDeptNames: string[];
      removeIds: Set<string>;
    }>();

    selectedDepts.forEach((term) => {
      const option = availableOptionMap[term.id];
      if (!option) {
        return;
      }

      const existingGroup = groupedChips.get(option.parentId) || {
        id: option.parentId,
        deptName: option.parentName,
        subDeptNames: [],
        removeIds: new Set<string>()
      };

      existingGroup.removeIds.add(term.id);

      if (option.indent === 1) {
        existingGroup.subDeptNames.push(option.term.name);
      }

      groupedChips.set(option.parentId, existingGroup);
    });

    return Array.from(groupedChips.values()).map((group) => ({
      id: group.id,
      label: group.subDeptNames.length > 0
        ? `${group.deptName} > ${group.subDeptNames.join(' | ')}`
        : group.deptName,
      removeIds: Array.from(group.removeIds)
    }));
  }, [availableOptionMap, selectedDepts]);

  return (
    <div ref={containerRef} style={containerStyle}>
      <label style={resolvedLabelStyle}>Department / Sub Department</label>
      <div style={{ position: 'relative' }}>
        <button
          type="button"
          style={resolvedTriggerStyle}
          onClick={() => {
            if (hasSelectedBUs) {
              setIsOpen((currentValue) => !currentValue);
            }
          }}
          disabled={!hasSelectedBUs}
        >
          {selectedDepts.length > 0 ? (
            <span style={chipWrapStyle}>
              {selectedChipItems.map((item) => (
                <span key={item.id} style={resolvedChipStyle}>
                  <span style={{ minWidth: 0, overflowWrap: 'anywhere' }}>{item.label}</span>
                  <button
                    type="button"
                    aria-label={`Remove ${item.label}`}
                    onClick={(event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      const removeIds = new Set(item.removeIds);
                      onChange(selectedDepts.filter((selectedTerm) => !removeIds.has(selectedTerm.id)));
                    }}
                    style={chipRemoveStyle}
                  >
                    ×
                  </button>
                </span>
              ))}
            </span>
          ) : (
            <span style={placeholderStyle}>{displayText}</span>
          )}
        </button>
        <span style={dropdownCaretStyle} />
        {selectedDepts.length > 1 && onClear && (
          <button
            type="button"
            aria-label="Clear Department / Sub Department"
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
      {isOpen && hasSelectedBUs && (
        <div style={menuStyle}>
          {!hasAvailableOptions ? (
            <div style={resolvedEmptyStateStyle}>No departments or sub departments available for this Business Unit</div>
          ) : availableOptions.map((option) => {
            const checked = selectedDepts.some((selectedTerm) => selectedTerm.id === option.term.id);

            return (
              <label
                key={option.term.id}
                style={{
                  ...resolvedOptionStyle,
                  paddingLeft: option.indent === 1 ? '28px' : '12px'
                }}
              >
                <input
                  type="checkbox"
                  checked={checked}
                  onChange={() => toggleSelection(option)}
                />
                <span>{option.indent === 1 ? `↳ ${option.term.name}` : option.term.name}</span>
              </label>
            );
          })}
        </div>
      )}
    </div>
  );
};

export default DeptSubDeptDropdown;

