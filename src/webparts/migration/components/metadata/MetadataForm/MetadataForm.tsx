import * as React from 'react';
import styles from './MetadataForm.module.scss';
import { IMetadataFormProps } from './IMetadataFormProps';
import { ITaxonomyTerm } from '../../../services/TaxonomyService';
import { TaxonomyPickerField } from '../fields/TaxonomyPickerField';
import { AuthorPickerField } from '../fields/AuthorPickerField';
import { fetchBUDepartmentTerms, Term } from '../../../utils/termStore';
import { BUDropdown } from '../fields/BUDropdown';
import { DeptSubDeptDropdown } from '../fields/DeptSubDeptDropdown';
import { CACHE_KEYS } from '../../../config/appConfig';

const MAX_SENSITIVE_TERMS_LENGTH = 1000;
const REVIEWER_STATUS_OPTIONS = ['Under Review', 'Active', 'Reject', 'Archive', 'Duplicate', 'Unpublish'];

const readAuthorCandidates = (value: any): string[] => {
  if (!value) {
    return [];
  }

  if (Array.isArray(value)) {
    return value.reduce<string[]>((authors, authorValue) => (
      authors.concat(readAuthorCandidates(authorValue))
    ), []);
  }

  if (Array.isArray(value.results)) {
    return readAuthorCandidates(value.results);
  }

  if (typeof value === 'object') {
    return [
      value.EMail,
      value.email,
      value.mail,
      value.upn,
      value.UserPrincipalName,
      value.userPrincipalName,
      value.Name,
      value.Key
    ].filter(Boolean).map(String);
  }

  return String(value)
    .split(/[;,]/)
    .map((part) => part.trim())
    .filter(Boolean);
};

const normalizeAuthorUpns = (
  sourceValues: Record<string, any> | undefined
): string[] => {
  const selectedCandidates = readAuthorCandidates(sourceValues?.selectedAuthorUpns)
    .concat(readAuthorCandidates(sourceValues?.selectedAuthorUpn));
  const authorCandidates = selectedCandidates.length > 0
    ? selectedCandidates
    : readAuthorCandidates(sourceValues?.author)
      .concat(readAuthorCandidates(sourceValues?.Author))
      .concat(readAuthorCandidates(sourceValues?.authorEmail))
      .concat(readAuthorCandidates(sourceValues?.authorUpn));
  const normalizedCandidates = authorCandidates
    .map((authorValue) => {
      const rawAuthorValue = String(authorValue || '').trim();
      return rawAuthorValue.indexOf('|') >= 0
        ? rawAuthorValue.split('|').pop() || rawAuthorValue
        : rawAuthorValue;
    })
    .filter(Boolean);
  const uniqueCandidates = Array.from(new Set(normalizedCandidates.map((authorValue) => authorValue.toLowerCase())))
    .map((lowerValue) => normalizedCandidates.find((authorValue) => authorValue.toLowerCase() === lowerValue) || lowerValue);

  return uniqueCandidates;
};

const buildInitialFormValues = (
  initialValues: Record<string, any> | undefined,
  context: IMetadataFormProps['context']
): Record<string, any> => ({
  title: '',
  description: '',
  buDepartmentTerms: [],
  buDepartmentTerm: null,
  geographyTerms: [],
  geographyTerm: null,
  clientTerms: [],
  clientTerm: null,
  documentTypeTerms: [],
  documentTypeTerm: null,
  diseaseAreaTerms: [],
  diseaseAreaTerm: null,
  therapyAreaTerms: [],
  therapyAreaTerm: null,
  sensitiveTerms: '',
  reviewerComments: '',
  reviewerCorrections: '',
  status: 'Under Review',
  ...initialValues,
  selectedAuthorUpns: normalizeAuthorUpns(initialValues)
});

export const MetadataForm: React.FC<IMetadataFormProps> = ({
  context,
  onSubmit,
  onClose,
  initialValues,
  taxonomyOptions,
  showReviewerFields = false,
  hideActions = false,
  density = 'default',
  onDraftChange
}) => {
  const [values, setValues] = React.useState<Record<string, any>>(() => buildInitialFormValues(initialValues, context));
  const [allTerms, setAllTerms] = React.useState<Term[]>([]);
  const [selectedBUs, setSelectedBUs] = React.useState<Term[]>([]);
  const [selectedDepts, setSelectedDepts] = React.useState<Term[]>([]);
  const [submitAttempted, setSubmitAttempted] = React.useState<boolean>(false);
  const [validationMessage, setValidationMessage] = React.useState<string>('');
  const [titleError, setTitleError] = React.useState('');
  const [isStatusOpen, setIsStatusOpen] = React.useState(false);
  const statusDropdownRef = React.useRef<HTMLDivElement | null>(null);
  const titleTextareaRef = React.useRef<HTMLTextAreaElement | null>(null);
  const descriptionTextareaRef = React.useRef<HTMLTextAreaElement | null>(null);
  const sensitiveTermsTextareaRef = React.useRef<HTMLTextAreaElement | null>(null);
  const reviewerCommentsTextareaRef = React.useRef<HTMLTextAreaElement | null>(null);
  const fieldRefs = React.useRef<Record<string, HTMLElement | null>>({});
  const onDraftChangeRef = React.useRef<IMetadataFormProps['onDraftChange']>(onDraftChange);

  // Track which fields were auto-filled (from initialValues)
  const autoFilledFields = React.useRef<Set<string>>(new Set());

  const resizeTextareaToContent = React.useCallback((textarea: HTMLTextAreaElement | null): void => {
    if (!textarea) {
      return;
    }

    textarea.classList.remove(styles.textareaHasOverflow);
    textarea.style.height = 'auto';
    const computedStyle = window.getComputedStyle(textarea);
    const maxHeight = parseFloat(computedStyle.maxHeight);
    const hasFiniteMaxHeight = Number.isFinite(maxHeight) && maxHeight > 0;
    const nextHeight = hasFiniteMaxHeight
      ? Math.min(textarea.scrollHeight, maxHeight)
      : textarea.scrollHeight;
    textarea.style.height = `${nextHeight}px`;

    if (textarea.scrollHeight > textarea.clientHeight + 1) {
      textarea.classList.add(styles.textareaHasOverflow);
    }
  }, []);

  React.useEffect(() => {
    if (initialValues) {
      if (initialValues.sensitiveTerms && initialValues.sensitiveTerms.trim()) {
        autoFilledFields.current.add('sensitiveTerms');
      }
      setValues(prev => ({
        ...prev,
        ...initialValues,
        selectedAuthorUpns: normalizeAuthorUpns(initialValues)
      }));
    }
  }, [initialValues]);

  React.useEffect(() => {
    const loadTerms = async (): Promise<void> => {
      try {
        const BU_KEY = CACHE_KEYS.buTerms;
        const cachedBU = (() => {
          try {
            const cached = sessionStorage.getItem(BU_KEY);
            return cached ? JSON.parse(cached) : null;
          } catch {
            return null;
          }
        })();
        if (cachedBU) {
          setAllTerms(cachedBU);
          return;
        }

        const terms = await fetchBUDepartmentTerms(context.spHttpClient, context.pageContext.web.absoluteUrl);
        setAllTerms(terms);
        try {
          sessionStorage.setItem(BU_KEY, JSON.stringify(terms));
        } catch {
          // Ignore session storage failures.
        }
      } catch (err) {
        console.error('Term fetch error:', err);
      }
    };

    void loadTerms();
  }, [context.pageContext.web.absoluteUrl, context.spHttpClient]);

  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent): void => {
      if (statusDropdownRef.current && !statusDropdownRef.current.contains(event.target as Node)) {
        setIsStatusOpen(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  React.useEffect(() => {
    resizeTextareaToContent(titleTextareaRef.current);
    resizeTextareaToContent(descriptionTextareaRef.current);
    resizeTextareaToContent(sensitiveTermsTextareaRef.current);
    resizeTextareaToContent(reviewerCommentsTextareaRef.current);
  }, [
    resizeTextareaToContent,
    values.description,
    values.reviewerComments,
    values.sensitiveTerms,
    values.title,
    showReviewerFields
  ]);

  React.useEffect(() => {
    onDraftChangeRef.current = onDraftChange;
  }, [onDraftChange]);

  const setFieldRef = React.useCallback((fieldName: string) => (node: HTMLElement | null): void => {
    fieldRefs.current[fieldName] = node;
  }, []);

  const scrollToFirstValidationError = React.useCallback((errors: Record<string, string>): void => {
    const orderedFieldNames = [
      'title',
      'bu',
      'author',
      'geography',
      'client',
      'documentType',
      'diseaseArea',
      'therapyArea',
      'description',
      'status'
    ];
    const firstErrorFieldName = orderedFieldNames.find((fieldName) => !!errors[fieldName]);
    const target = firstErrorFieldName ? fieldRefs.current[firstErrorFieldName] : null;

    if (!target) {
      return;
    }

    window.requestAnimationFrame(() => {
      target.scrollIntoView({ behavior: 'smooth', block: 'center' });
      const focusTarget = target.querySelector<HTMLElement>(
        'textarea, input, select, button, [tabindex]:not([tabindex="-1"])'
      );
      focusTarget?.focus?.({ preventScroll: true });
    });
  }, []);

  React.useEffect(() => {
    const selectedTerms = Array.isArray(values.buDepartmentTerms) ? values.buDepartmentTerms.filter(Boolean) : [];
    if (allTerms.length === 0 || selectedTerms.length === 0) {
      return;
    }

    const nextSelectedBUs: Term[] = [];
    const nextSelectedDepts: Term[] = [];
    const seenBuIds = new Set<string>();
    const seenDeptIds = new Set<string>();

    selectedTerms.forEach((selectedTerm) => {
      const matchedRoot = allTerms.find((bu) => bu.id === selectedTerm.id);
      if (matchedRoot) {
        if (!seenBuIds.has(matchedRoot.id)) {
          seenBuIds.add(matchedRoot.id);
          nextSelectedBUs.push(matchedRoot);
        }
        return;
      }

      allTerms.forEach((bu) => {
        const matchedDept = (bu.children || []).find((dept) => dept.id === selectedTerm.id);
        if (matchedDept) {
          if (!seenBuIds.has(bu.id)) {
            seenBuIds.add(bu.id);
            nextSelectedBUs.push(bu);
          }
          if (!seenDeptIds.has(matchedDept.id)) {
            seenDeptIds.add(matchedDept.id);
            nextSelectedDepts.push(matchedDept);
          }
          return;
        }

        (bu.children || []).forEach((dept) => {
          const matchedSubDept = (dept.children || []).find((subDept) => subDept.id === selectedTerm.id);
          if (matchedSubDept) {
            if (!seenBuIds.has(bu.id)) {
              seenBuIds.add(bu.id);
              nextSelectedBUs.push(bu);
            }
            if (!seenDeptIds.has(matchedSubDept.id)) {
              seenDeptIds.add(matchedSubDept.id);
              nextSelectedDepts.push(matchedSubDept);
            }
          }
        });
      });
    });

    setSelectedBUs(nextSelectedBUs);
    setSelectedDepts(nextSelectedDepts);
  }, [allTerms, values.buDepartmentTerms]);

  // Only Sensitive Terms should trigger the sensitive-info warning.
  const hasSensitiveInfo = autoFilledFields.current.has('sensitiveTerms') && !!values.sensitiveTerms;

  const updateFieldValue = React.useCallback((name: string, value: string) => {
    setValues((currentValues) => ({
      ...currentValues,
      [name]: value
    }));
  }, []);

  const onChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    if (e.target instanceof HTMLTextAreaElement) {
      resizeTextareaToContent(e.target);
    }
    updateFieldValue(name, value);
  };

  const validateTitle = (value: string): string => {
    const specialChars = /[$%@*+?!#&^~`|\\<>]/;
    if (specialChars.test(value)) {
      return 'Special characters ($, %, @, *, +, ?, !, #) are not allowed';
    }
    if (value.length > 255) {
      return 'Title must be 255 characters or less';
    }
    return '';
  };

  const handleTitleChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    const val = e.target.value;
    const sanitizedValue = val.replace(/[$%@*+?!#&^~`|\\<>]/g, '');

    setTitleError(validateTitle(val));
    resizeTextareaToContent(e.target);
    updateFieldValue('title', sanitizedValue);
  };

  const onTaxonomyChange = (fieldName: string, term: ITaxonomyTerm | null) => {
    setValues((currentValues) => ({
      ...currentValues,
      [fieldName]: term
    }));
  };

  const onMultiTaxonomyChange = (fieldName: string, singularFieldName: string, terms: ITaxonomyTerm[]) => {
    setValues((currentValues) => ({
      ...currentValues,
      [fieldName]: terms,
      [singularFieldName]: terms[0] || null
    }));
  };

  const clearTextField = (fieldName: string) => {
    if (fieldName === 'title') {
      setTitleError('');
    }

    updateFieldValue(fieldName, '');
  };

  const validationState = React.useMemo(() => {
    const errors: Record<string, string> = {};
    const titleValue = String(values.title || '');
    const titleValidationMessage = validateTitle(titleValue);

    if (!titleValue.trim()) {
      errors.title = 'Title is required';
    } else if (titleValidationMessage) {
      errors.title = titleValidationMessage;
    }

    if (selectedBUs.length === 0) {
      errors.bu = 'Business Unit is required';
    }

    if (!Array.isArray(values.selectedAuthorUpns) || values.selectedAuthorUpns.length === 0) {
      errors.author = 'Author is required';
    }

    if (!Array.isArray(values.geographyTerms) || values.geographyTerms.length === 0) {
      errors.geography = 'Geography is required';
    }

    if (!Array.isArray(values.clientTerms) || values.clientTerms.length === 0) {
      errors.client = 'Client is required';
    }

    if (!Array.isArray(values.documentTypeTerms) || values.documentTypeTerms.length === 0) {
      errors.documentType = 'Document Type is required';
    }

    if (!Array.isArray(values.diseaseAreaTerms) || values.diseaseAreaTerms.length === 0) {
      errors.diseaseArea = 'Disease Area is required';
    }

    if (!Array.isArray(values.therapyAreaTerms) || values.therapyAreaTerms.length === 0) {
      errors.therapyArea = 'Therapy Area is required';
    }

    if (!String(values.description || '').trim()) {
      errors.description = 'Description is required';
    }

    if (showReviewerFields) {
      if (!String(values.status || '').trim()) {
        errors.status = 'Status is required';
      }
    }

    return {
      errors,
      isValid: Object.keys(errors).length === 0
    };
  }, [
    selectedBUs.length,
    values.clientTerms,
    values.description,
    values.diseaseAreaTerms,
    values.documentTypeTerms,
    values.geographyTerms,
    values.reviewerComments,
    values.selectedAuthorUpns,
    values.status,
    values.therapyAreaTerms,
    values.title,
    showReviewerFields
  ]);

  const normalizedBuDepartmentTerms = React.useMemo(() => (
    selectedDepts.length > 0
      ? selectedDepts
      : selectedBUs.length > 0
        ? selectedBUs
        : Array.isArray(values.buDepartmentTerms)
          ? values.buDepartmentTerms
          : []
  ), [selectedBUs, selectedDepts, values.buDepartmentTerms]);

  const submittedValues = React.useMemo(() => ({
    ...values,
    sensitiveTerms: String(values.sensitiveTerms || '').slice(0, MAX_SENSITIVE_TERMS_LENGTH),
    buDepartmentTerms: normalizedBuDepartmentTerms,
    buDepartmentTerm: normalizedBuDepartmentTerms[0] || null,
    selectedBUs,
    selectedDepts
  }), [normalizedBuDepartmentTerms, selectedBUs, selectedDepts, values]);

  React.useEffect(() => {
    onDraftChangeRef.current?.(submittedValues, validationState.isValid);
  }, [submittedValues, validationState.isValid]);

  const handleSubmit = (e?: React.FormEvent) => {
    e && e.preventDefault();
    setSubmitAttempted(true);

    if (!validationState.isValid) {
      setValidationMessage('Please fill all mandatory fields before submitting.');
      scrollToFirstValidationError(validationState.errors);
      return;
    }

    setValidationMessage('');
    onSubmit && onSubmit(submittedValues);
  };

  const renderRequiredLabel = React.useCallback((text: string, isRequired: boolean = false) => (
    <span className={styles.labelText}>
      {text}
      {isRequired && <span className={styles.requiredMark}>*</span>}
    </span>
  ), []);

  const renderFieldError = React.useCallback((fieldName: string): JSX.Element | null => {
    const fieldMessage = validationState.errors[fieldName];

    if (!submitAttempted || !fieldMessage) {
      return null;
    }

    return (
      <div className={styles.fieldValidationMessage} role="alert">
        {fieldMessage}
      </div>
    );
  }, [submitAttempted, validationState.errors]);

  return (
    <div className={`${styles.formWrapper} ${density === 'compact' ? styles.compactDensity : ''}`}>
      <div className={styles.headerBar}>
        <div className={styles.titleWrap}>
          <h2 className={styles.title}>Document Metadata</h2>
          <div className={styles.subtitle}>AI auto-generates title, description, classification and tags. You can review and edit these as needed, and update author, business unit and department details (if applicable) before submission.</div>
          {hasSensitiveInfo && (
            <div className={styles.sensitiveWarning}>
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="2" />
                <line x1="12" y1="8" x2="12" y2="12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                <line x1="12" y1="16" x2="12.01" y2="16" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
              </svg>
              <span>This file contains sensitive information</span>
            </div>
          )}
        </div>
      </div>

      <div className={styles.cardMeta}>
        <form onSubmit={handleSubmit}>
          <div className={styles.formInner}>
            {submitAttempted && validationMessage && (
              <div className={styles.validationNotice} role="alert">
                {validationMessage}
              </div>
            )}
            <div className={styles.grid}>
              <div ref={setFieldRef('title')} className={styles.fieldFull}>
                <div className={styles.labelRow}>
                  <label className={styles.label} htmlFor="title">{renderRequiredLabel('Title', true)}</label>
                  {String(values.title || '').trim() && (
                    <button
                      type="button"
                      className={styles.clearValueBtn}
                      aria-label="Clear Title"
                      onClick={() => clearTextField('title')}
                    >
                      ×
                    </button>
                  )}
                </div>
                <textarea
                  ref={titleTextareaRef}
                  id="title"
                  name="title"
                  rows={1}
                  placeholder="Value goes here"
                  value={values.title}
                  onChange={handleTitleChange}
                  maxLength={255}
                  className={`${styles.titleTextarea} ${submitAttempted && validationState.errors.title ? styles.fieldError : ''}`}
                />
                {titleError && (
                  <span style={{
                    color: '#a4262c',
                    fontSize: '12px',
                    marginTop: '4px',
                    display: 'block'
                  }}>
                    {titleError}
                  </span>
                )}
                {renderFieldError('title')}
              </div>

              <div ref={setFieldRef('bu')} className={styles.field}>
                <div className={submitAttempted && validationState.errors.bu ? styles.fieldErrorWrap : undefined}>
                  <BUDropdown
                    buTerms={allTerms}
                    selectedBUs={selectedBUs}
                    onChange={setSelectedBUs}
                    onClear={() => setSelectedBUs([])}
                    required={true}
                    density={density}
                  />
                </div>
                {renderFieldError('bu')}
              </div>

              <div ref={setFieldRef('author')} className={styles.field}>
                <DeptSubDeptDropdown
                  buTerms={allTerms}
                  selectedBUs={selectedBUs}
                  selectedDepts={selectedDepts}
                  onChange={setSelectedDepts}
                  onClear={() => setSelectedDepts([])}
                  density={density}
                />
              </div>

              <div ref={setFieldRef('geography')} className={styles.field}>
                <label className={styles.label} htmlFor="author">{renderRequiredLabel('Author', true)}</label>
                <AuthorPickerField
                  context={context}
                  selectedAuthorUpns={values.selectedAuthorUpns || []}
                  onChange={(upns) => setValues((currentValues) => ({
                    ...currentValues,
                    selectedAuthorUpns: upns
                  }))}
                  placeholder="Search KMS Users"
                  className={`${styles.select} ${submitAttempted && validationState.errors.author ? styles.fieldError : ''}`}
                  density={density}
                />
                {renderFieldError('author')}
              </div>

              <div ref={setFieldRef('client')} className={styles.field}>
                <label className={styles.label} htmlFor="geography">{renderRequiredLabel('Geography', true)}</label>
                <TaxonomyPickerField
                  id="geography"
                  options={taxonomyOptions.geography}
                  placeholder="Value goes here"
                  multiSelect={true}
                  value={values.geographyTerms || []}
                  onChange={(terms) => onMultiTaxonomyChange('geographyTerms', 'geographyTerm', (terms as ITaxonomyTerm[]) || [])}
                  className={`${styles.select} ${submitAttempted && validationState.errors.geography ? styles.fieldError : ''}`}
                />
                {renderFieldError('geography')}
              </div>

              <div ref={setFieldRef('documentType')} className={styles.field}>
                <label className={styles.label} htmlFor="client">{renderRequiredLabel('Client', true)}</label>
                <TaxonomyPickerField
                  id="client"
                  options={taxonomyOptions.client}
                  placeholder="Value goes here"
                  multiSelect={true}
                  value={values.clientTerms || []}
                  onChange={(terms) => onMultiTaxonomyChange('clientTerms', 'clientTerm', (terms as ITaxonomyTerm[]) || [])}
                  className={`${styles.select} ${submitAttempted && validationState.errors.client ? styles.fieldError : ''}`}
                />
                {renderFieldError('client')}
              </div>

              <div ref={setFieldRef('diseaseArea')} className={styles.field}>
                <label className={styles.label} htmlFor="documentType">{renderRequiredLabel('Document Type', true)}</label>
                <TaxonomyPickerField
                  id="documentType"
                  options={taxonomyOptions.documentType}
                  placeholder="Value goes here"
                  multiSelect={true}
                  value={values.documentTypeTerms || []}
                  onChange={(terms) => onMultiTaxonomyChange('documentTypeTerms', 'documentTypeTerm', (terms as ITaxonomyTerm[]) || [])}
                  className={`${styles.select} ${submitAttempted && validationState.errors.documentType ? styles.fieldError : ''}`}
                />
                {renderFieldError('documentType')}
              </div>

              <div ref={setFieldRef('therapyArea')} className={styles.field}>
                <label className={styles.label} htmlFor="diseaseArea">{renderRequiredLabel('Disease Area', true)}</label>
                <TaxonomyPickerField
                  id="diseaseArea"
                  options={taxonomyOptions.diseaseArea}
                  placeholder="Value goes here"
                  multiSelect={true}
                  value={values.diseaseAreaTerms || []}
                  onChange={(terms) => onMultiTaxonomyChange('diseaseAreaTerms', 'diseaseAreaTerm', (terms as ITaxonomyTerm[]) || [])}
                  className={`${styles.select} ${submitAttempted && validationState.errors.diseaseArea ? styles.fieldError : ''}`}
                />
                {renderFieldError('diseaseArea')}
              </div>

              <div className={styles.field}>
                <label className={styles.label} htmlFor="therapyArea">{renderRequiredLabel('Therapy Area', true)}</label>
                <TaxonomyPickerField
                  id="therapyArea"
                  options={taxonomyOptions.therapyArea}
                  placeholder="Value goes here"
                  multiSelect={true}
                  value={values.therapyAreaTerms || []}
                  onChange={(terms) => onMultiTaxonomyChange('therapyAreaTerms', 'therapyAreaTerm', (terms as ITaxonomyTerm[]) || [])}
                  className={`${styles.select} ${submitAttempted && validationState.errors.therapyArea ? styles.fieldError : ''}`}
                />
                {renderFieldError('therapyArea')}
              </div>

              <div ref={setFieldRef('description')} className={styles.fieldFull}>
                <div className={styles.labelRow}>
                  <label className={styles.label} htmlFor="description">{renderRequiredLabel('Description', true)}</label>
                  {String(values.description || '').trim() && (
                    <button
                      type="button"
                      className={styles.clearValueBtn}
                      aria-label="Clear Description"
                      onClick={() => clearTextField('description')}
                    >
                      ×
                    </button>
                  )}
                </div>
                <textarea
                  ref={descriptionTextareaRef}
                  id="description"
                  name="description"
                  placeholder="Vid elementum eros ullamcorper id. Duis tempor fermentum tortor, ac fermentum orci varius eget. Vestibulum lorem arcu, placerat et suscipit ut"
                  value={values.description}
                  onChange={onChange}
                  className={`${styles.descriptionTextarea} ${submitAttempted && validationState.errors.description ? styles.fieldError : ''}`}
                />
                {renderFieldError('description')}
              </div>

              <div className={styles.fieldFull}>
                <div className={styles.labelRow}>
                  <label className={styles.label} htmlFor="sensitiveTerms">{renderRequiredLabel('Sensitive information')}</label>
                  {String(values.sensitiveTerms || '').trim() && (
                    <button
                      type="button"
                      className={styles.clearValueBtn}
                      aria-label="Clear Sensitive Terms"
                      onClick={() => clearTextField('sensitiveTerms')}
                    >
                      ×
                    </button>
                  )}
                </div>
                <textarea
                  ref={sensitiveTermsTextareaRef}
                  id="sensitiveTerms"
                  name="sensitiveTerms"
                  placeholder="Value goes here"
                  value={values.sensitiveTerms}
                  onChange={onChange}
                  maxLength={MAX_SENSITIVE_TERMS_LENGTH}
                  className={`${styles.textareaScrollable} ${autoFilledFields.current.has('sensitiveTerms') && values.sensitiveTerms ? styles.inputSensitive : ''}`}
                />
              </div>

              {showReviewerFields && (
                <>
                  <div ref={setFieldRef('status')} className={`${styles.field} ${styles.reviewerStatusField}`}>
                    <label className={styles.label} htmlFor="status">{renderRequiredLabel('Status', true)}</label>
                    <div ref={statusDropdownRef} className={styles.statusDropdown}>
                      <button
                        id="status"
                        type="button"
                        className={`${styles.statusDropdownTrigger} ${submitAttempted && validationState.errors.status ? styles.fieldError : ''}`}
                        aria-haspopup="listbox"
                        aria-expanded={isStatusOpen}
                        onClick={() => setIsStatusOpen((currentValue) => !currentValue)}
                      >
                        <span>{values.status || 'Under Review'}</span>
                        <span className={styles.statusDropdownCaret} aria-hidden="true" />
                      </button>
                      {isStatusOpen && (
                        <div className={styles.statusDropdownMenu} role="listbox" aria-label="Status">
                      {Array.from(new Set([String(values.status || 'Under Review'), ...REVIEWER_STATUS_OPTIONS])).map((option) => (
                            <button
                              key={option}
                              type="button"
                              role="option"
                              aria-selected={(values.status || 'Under Review') === option}
                              className={`${styles.statusDropdownOption} ${(values.status || 'Under Review') === option ? styles.statusDropdownOptionSelected : ''}`}
                              onClick={() => {
                                updateFieldValue('status', option);
                                setIsStatusOpen(false);
                              }}
                            >
                              {option}
                            </button>
                      ))}
                        </div>
                      )}
                    </div>
                    {renderFieldError('status')}
                  </div>

                  <div className={styles.fieldFull}>
                    <label className={styles.label} htmlFor="reviewerComments">Reviewer Comments</label>
                    <textarea
                      ref={reviewerCommentsTextareaRef}
                      id="reviewerComments"
                      name="reviewerComments"
                      placeholder="Add reviewer comments"
                      value={values.reviewerComments || ''}
                      onChange={onChange}
                      maxLength={500}
                      className={styles.textarea}
                    />
                    <span className={styles.charCount}>{String(values.reviewerComments || '').length}/500</span>
                  </div>
                </>
              )}

            </div>

            {!hideActions && (
              <div className={styles.actions}>
                <button type="submit" className={styles.submitBtn}>Submit</button>
                <button type="button" className={styles.closeActionBtn} onClick={() => onClose && onClose()}>Close</button>
              </div>
            )}
          </div>
        </form>
      </div>
    </div>
  );

};

export default MetadataForm;

