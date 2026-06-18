import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { IRenameBusinessUnitDialogProps } from './IRenameBusinessUnitDialogProps';
import { BUSINESS_UNIT_HIERARCHY } from './BusinessUnitHierarchy';
import { LIBRARY_NAMES, PAGE_SIZES } from '../../../config/appConfig';
import styles from './RenameBusinessUnitDialog.module.scss';

const LIBRARY_NAME = LIBRARY_NAMES.kmDataHub;
const UPDATE_BATCH_SIZE = PAGE_SIZES.updateBatchSize;
const NONE_OPTION = 'None';

interface IHierarchyItem {
  Id: number;
  BusinessUnit: string;
  Department: string;
  [key: string]: any;
}

interface ISharePointField {
  InternalName: string;
  Title: string;
  Hidden?: boolean;
  ReadOnlyField?: boolean;
}

const normalizeValue = (value?: string): string => (value || '').trim();

const isMeaningfulValue = (value?: string): boolean => {
  const normalized = normalizeValue(value);
  return normalized.length > 0 && normalized !== '-';
};

const toUniqueSortedValues = (values: string[]): string[] =>
  Array.from(new Set(values.map(normalizeValue).filter(isMeaningfulValue))).sort((left, right) => left.localeCompare(right));

const normalizeFieldKey = (value?: string): string => normalizeValue(value).toLowerCase().replace(/[^a-z0-9]/g, '');

const toHierarchyOptions = (values: string[]): string[] => {
  const normalized = toUniqueSortedValues(values);
  return normalized.length > 0 ? normalized : [NONE_OPTION];
};

export const RenameBusinessUnitDialog: React.FC<IRenameBusinessUnitDialogProps> = ({ context, isOpen, onClose, variant = 'modal' }) => {
  const [loading, setLoading] = React.useState(false);
  const [submitting, setSubmitting] = React.useState(false);
  const [items, setItems] = React.useState<IHierarchyItem[]>([]);
  const [subDepartmentFieldName, setSubDepartmentFieldName] = React.useState<string | null>(null);
  const [entityTypeName, setEntityTypeName] = React.useState<string>('');

  const [selectedBU, setSelectedBU] = React.useState('');
  const [targetBU, setTargetBU] = React.useState('');

  const [renameDepartment, setRenameDepartment] = React.useState(false);
  const [selectedDepartment, setSelectedDepartment] = React.useState('');
  const [targetDepartment, setTargetDepartment] = React.useState('');

  const [renameSubDepartment, setRenameSubDepartment] = React.useState(false);
  const [selectedSubDepartment, setSelectedSubDepartment] = React.useState('');
  const [targetSubDepartment, setTargetSubDepartment] = React.useState('');

  const [errorMessage, setErrorMessage] = React.useState<string | null>(null);
  const [successMessage, setSuccessMessage] = React.useState<string | null>(null);
  const [confirmationCount, setConfirmationCount] = React.useState<number | null>(null);

  const webUrl = context.pageContext.web.absoluteUrl;

  const resetForm = React.useCallback(() => {
    setSelectedBU('');
    setTargetBU('');
    setRenameDepartment(false);
    setSelectedDepartment('');
    setTargetDepartment('');
    setRenameSubDepartment(false);
    setSelectedSubDepartment('');
    setTargetSubDepartment('');
    setErrorMessage(null);
    setSuccessMessage(null);
    setConfirmationCount(null);
  }, []);

  const getListInfo = React.useCallback(async (): Promise<string> => {
    const response = await context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')?$select=ListItemEntityTypeFullName`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to load ${LIBRARY_NAME} metadata.`);
    }

    const data = await response.json();
    return data.ListItemEntityTypeFullName;
  }, [context.spHttpClient, webUrl]);

  const createSubDepartmentField = React.useCallback(async (): Promise<string> => {
    const response = await context.spHttpClient.post(
      `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/fields/createfieldasxml`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'odata-version': ''
        },
        body: JSON.stringify({
          parameters: {
            SchemaXml: `<Field Type="Text" DisplayName="Sub-Department" Name="SubDepartment" StaticName="SubDepartment" Group="Custom Columns" />`,
            Options: 0
          }
        })
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to create Sub-Department column: ${response.status} ${errorText}`);
    }

    return 'SubDepartment';
  }, [context.spHttpClient, webUrl]);

  const detectSubDepartmentField = React.useCallback(async (): Promise<string | null> => {
    const response = await context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/fields?$select=InternalName,Title,Hidden,ReadOnlyField`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      return null;
    }

    const data = await response.json();
    const fields: ISharePointField[] = data.value || [];

    const exactMatch = fields.find((field) => {
      const candidates = [field.InternalName, field.Title].map(normalizeFieldKey);
      return candidates.indexOf('subdepartment') !== -1;
    });

    if (exactMatch) {
      return exactMatch.InternalName;
    }

    const fuzzyMatch = fields.find((field) => {
      const candidates = [field.InternalName, field.Title].map(normalizeFieldKey);
      return candidates.some((candidate) => candidate.indexOf('sub') !== -1 && candidate.indexOf('department') !== -1);
    });

    return fuzzyMatch ? fuzzyMatch.InternalName : null;
  }, [context.spHttpClient, webUrl]);

  const ensureSubDepartmentField = React.useCallback(async (): Promise<string> => {
    const detectedField = await detectSubDepartmentField();
    if (detectedField) {
      return detectedField;
    }

    return createSubDepartmentField();
  }, [createSubDepartmentField, detectSubDepartmentField]);

  const fetchAllHierarchyItems = React.useCallback(async (subFieldName: string | null): Promise<IHierarchyItem[]> => {
    const selectFields = ['Id', 'BusinessUnit', 'Department'];
    if (subFieldName) {
      selectFields.push(subFieldName);
    }

    let requestUrl =
      `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items` +
      `?$select=${selectFields.join(',')}` +
      `&$orderby=Id asc&$top=5000`;

    const collectedItems: IHierarchyItem[] = [];

    while (requestUrl) {
      const response = await context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);

      if (!response.ok) {
        throw new Error(`Failed to load ${LIBRARY_NAME} hierarchy data.`);
      }

      const data = await response.json();
      const pageItems: IHierarchyItem[] = (data.value || []).map((item: any) => ({
        Id: item.Id,
        BusinessUnit: item.BusinessUnit || '',
        Department: item.Department || '',
        ...(subFieldName ? { [subFieldName]: item[subFieldName] || '' } : {})
      }));

      collectedItems.push(...pageItems);
      requestUrl = data['@odata.nextLink'] || data.__next || '';
    }

    return collectedItems;
  }, [context.spHttpClient, webUrl]);

  const updateItemWithEntityType = React.useCallback(async (
    itemId: number,
    payload: Record<string, string>,
    entityType: string
  ) => {
    const response = await context.spHttpClient.post(
      `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE',
          'odata-version': ''
        },
        body: JSON.stringify({
          __metadata: { type: entityType },
          ...payload
        })
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to update item ${itemId}: ${response.status} ${errorText}`);
    }
  }, [context.spHttpClient, webUrl]);

  const backfillMissingHierarchyValues = React.useCallback(async (
    hierarchyItems: IHierarchyItem[],
    entityType: string,
    subFieldName: string
  ): Promise<void> => {
    const itemsNeedingBackfill = hierarchyItems
      .map((item) => {
        const payload: Record<string, string> = {};

        if (!isMeaningfulValue(item.Department)) {
          payload.Department = NONE_OPTION;
        }

        if (!isMeaningfulValue(item[subFieldName])) {
          payload[subFieldName] = NONE_OPTION;
        }

        return {
          itemId: item.Id,
          payload
        };
      })
      .filter((entry) => Object.keys(entry.payload).length > 0);

    for (let index = 0; index < itemsNeedingBackfill.length; index += UPDATE_BATCH_SIZE) {
      const batch = itemsNeedingBackfill.slice(index, index + UPDATE_BATCH_SIZE);

      await Promise.all(
        batch.map((entry) => updateItemWithEntityType(entry.itemId, entry.payload, entityType))
      );

      if (index + UPDATE_BATCH_SIZE < itemsNeedingBackfill.length) {
        await new Promise((resolve) => setTimeout(resolve, 150));
      }
    }
  }, [updateItemWithEntityType]);

  const loadHierarchy = React.useCallback(async () => {
    setLoading(true);
    setErrorMessage(null);

    try {
      const [entityType, ensuredSubDepartmentField] = await Promise.all([
        getListInfo(),
        ensureSubDepartmentField()
      ]);

      let hierarchyItems = await fetchAllHierarchyItems(ensuredSubDepartmentField);
      await backfillMissingHierarchyValues(hierarchyItems, entityType, ensuredSubDepartmentField);
      hierarchyItems = await fetchAllHierarchyItems(ensuredSubDepartmentField);

      setEntityTypeName(entityType);
      setSubDepartmentFieldName(ensuredSubDepartmentField);
      setItems(hierarchyItems);
    } catch (error) {
      setErrorMessage(error instanceof Error ? error.message : 'Unable to load Business Unit hierarchy.');
    } finally {
      setLoading(false);
    }
  }, [backfillMissingHierarchyValues, ensureSubDepartmentField, fetchAllHierarchyItems, getListInfo]);

  React.useEffect(() => {
    if (!isOpen) {
      return;
    }

    resetForm();
    loadHierarchy().catch(() => undefined);
  }, [isOpen, loadHierarchy, resetForm]);

  React.useEffect(() => {
    setRenameDepartment(false);
    setRenameSubDepartment(false);
    setSelectedDepartment('');
    setTargetDepartment('');
    setSelectedSubDepartment('');
    setTargetSubDepartment('');
    setConfirmationCount(null);
    setErrorMessage(null);
    setSuccessMessage(null);
  }, [selectedBU, targetBU]);

  React.useEffect(() => {
    if (!renameDepartment) {
      setRenameSubDepartment(false);
      setSelectedDepartment('');
      setTargetDepartment('');
      setSelectedSubDepartment('');
      setTargetSubDepartment('');
    }

    setConfirmationCount(null);
    setErrorMessage(null);
    setSuccessMessage(null);
  }, [renameDepartment]);

  React.useEffect(() => {
    if (!renameSubDepartment) {
      setSelectedSubDepartment('');
      setTargetSubDepartment('');
    }

    setConfirmationCount(null);
    setErrorMessage(null);
  }, [renameSubDepartment]);

  React.useEffect(() => {
    setSelectedSubDepartment('');
    setTargetSubDepartment('');
    setConfirmationCount(null);
    setErrorMessage(null);
  }, [selectedDepartment, targetDepartment]);

  const businessUnitOptions = React.useMemo(
    () => Object.keys(BUSINESS_UNIT_HIERARCHY).sort((left, right) => left.localeCompare(right)),
    []
  );

  const departmentOptions = React.useMemo(
    () => toHierarchyOptions(Object.keys(BUSINESS_UNIT_HIERARCHY[selectedBU] || {})),
    [selectedBU]
  );

  const targetDepartmentOptions = React.useMemo(
    () => toHierarchyOptions(Object.keys(BUSINESS_UNIT_HIERARCHY[targetBU] || {})),
    [targetBU]
  );

  const subDepartmentOptions = React.useMemo(
    () => toHierarchyOptions(((BUSINESS_UNIT_HIERARCHY[selectedBU] || {})[selectedDepartment] || []).slice()),
    [selectedBU, selectedDepartment]
  );

  const targetSubDepartmentOptions = React.useMemo(
    () => toHierarchyOptions(((BUSINESS_UNIT_HIERARCHY[targetBU] || {})[targetDepartment] || []).slice()),
    [targetBU, targetDepartment]
  );

  const selectedBUItems = React.useMemo(
    () => items.filter((item) => normalizeValue(item.BusinessUnit) === normalizeValue(selectedBU)),
    [items, selectedBU]
  );

  const hasCompletedBusinessUnitStep = Boolean(selectedBU && targetBU);
  const hasCompletedDepartmentStep = Boolean(selectedDepartment && targetDepartment);

  const validate = React.useCallback((): string | null => {
    if (!selectedBU) {
      return 'Please select an existing Business Unit.';
    }

    if (!targetBU) {
      return 'Please select the target Business Unit.';
    }

    if (renameDepartment) {
      if (!selectedDepartment) {
        return 'Please select the Department to rename.';
      }

      if (!targetDepartment) {
        return 'Please select the target Department.';
      }
    }

    if (renameSubDepartment) {
      if (!selectedDepartment) {
        return 'Please select a Department before renaming a Sub-Department.';
      }

      if (!subDepartmentFieldName) {
        return 'Sub-Department field was not found in the KM Artifacts library.';
      }

      if (!selectedSubDepartment) {
        return 'Please select the Sub-Department to rename.';
      }

      if (!targetDepartment) {
        return 'Please select the target Department before selecting a target Sub-Department.';
      }

      if (!targetSubDepartment) {
        return 'Please select the target Sub-Department.';
      }
    }

    return null;
  }, [
    renameDepartment,
    renameSubDepartment,
    selectedBU,
    selectedDepartment,
    selectedSubDepartment,
    subDepartmentFieldName,
    targetBU,
    targetDepartment,
    targetSubDepartment
  ]);

  const handlePrepareUpdate = async (): Promise<void> => {
    const validationError = validate();
    if (validationError) {
      setErrorMessage(validationError);
      setConfirmationCount(null);
      return;
    }

    if (selectedBUItems.length === 0) {
      setErrorMessage('No documents were found for the selected Business Unit.');
      setConfirmationCount(null);
      return;
    }

    setErrorMessage(null);
    setSuccessMessage(null);
    setConfirmationCount(selectedBUItems.length);
  };

  const updateItem = React.useCallback(async (itemId: number, payload: Record<string, string>) => {
    await updateItemWithEntityType(itemId, payload, entityTypeName);
  }, [entityTypeName, updateItemWithEntityType]);

  const handleConfirmUpdate = async (): Promise<void> => {
    if (!entityTypeName) {
      setErrorMessage('List metadata is not ready. Please try again.');
      return;
    }

    setSubmitting(true);
    setErrorMessage(null);

    try {
      const itemsToUpdate = selectedBUItems.map((item) => {
        const payload: Record<string, string> = {
          BusinessUnit: targetBU
        };

        if (renameDepartment && normalizeValue(item.Department) === normalizeValue(selectedDepartment)) {
          payload.Department = targetDepartment;
        }

        if (
          renameSubDepartment &&
          subDepartmentFieldName &&
          normalizeValue(item.Department) === normalizeValue(selectedDepartment) &&
          normalizeValue(item[subDepartmentFieldName]) === normalizeValue(selectedSubDepartment)
        ) {
          payload[subDepartmentFieldName] = targetSubDepartment;
        }

        return {
          itemId: item.Id,
          payload
        };
      });

      for (let index = 0; index < itemsToUpdate.length; index += UPDATE_BATCH_SIZE) {
        const batch = itemsToUpdate.slice(index, index + UPDATE_BATCH_SIZE);

        await Promise.all(
          batch.map((entry) => updateItem(entry.itemId, entry.payload))
        );

        if (index + UPDATE_BATCH_SIZE < itemsToUpdate.length) {
          await new Promise((resolve) => setTimeout(resolve, 150));
        }
      }

      setSuccessMessage(`Successfully updated ${itemsToUpdate.length} documents.`);
      setConfirmationCount(null);
      await loadHierarchy();
      resetForm();
    } catch (error) {
      setErrorMessage(error instanceof Error ? error.message : 'Failed to update the selected documents.');
    } finally {
      setSubmitting(false);
    }
  };

  const handleClose = (): void => {
    if (submitting) {
      return;
    }

    resetForm();
    onClose();
  };

  if (!isOpen) {
    return null;
  }

  const dialogContent = (
      <div className={`${styles.dialog} ${variant === 'inline' ? styles.inlineDialog : ''}`}>
        <div className={styles.header}>
          <div>
            <h2 className={styles.title}>Rename BU</h2>
            <p className={styles.subtitle}>Map existing hierarchy values to the approved BU, Department, and Sub-Department combinations you shared.</p>
          </div>
          <button type="button" className={styles.closeButton} onClick={handleClose} aria-label={variant === 'inline' ? 'Back to Business Unit tools' : 'Close Rename BU dialog'}>
            {variant === 'inline' ? '←' : '×'}
          </button>
        </div>

        <div className={styles.body}>
          {loading ? (
            <div className={styles.stateCard}>Loading hierarchy...</div>
          ) : (
            <>
              {errorMessage && <div className={`${styles.message} ${styles.errorMessage}`}>{errorMessage}</div>}
              {successMessage && <div className={`${styles.message} ${styles.successMessage}`}>{successMessage}</div>}

              <div className={styles.card}>
                <h3 className={styles.cardTitle}>Business Unit</h3>
                <div className={styles.fieldGrid}>
                  <label className={styles.field}>
                    <span className={styles.label}>Select existing BU</span>
                    <select value={selectedBU} onChange={(event) => setSelectedBU(event.target.value)} className={styles.select}>
                      <option value="">Select Business Unit</option>
                      {businessUnitOptions.map((option) => (
                        <option key={option} value={option}>{option}</option>
                      ))}
                    </select>
                  </label>

                  <label className={styles.field}>
                    <span className={styles.label}>Select target BU</span>
                    <select value={targetBU} onChange={(event) => setTargetBU(event.target.value)} className={styles.select}>
                      <option value="">Select target Business Unit</option>
                      {businessUnitOptions.map((option) => (
                        <option key={option} value={option}>{option}</option>
                      ))}
                    </select>
                  </label>
                </div>
              </div>

              {hasCompletedBusinessUnitStep && (
                <div className={styles.card}>
                  <label className={styles.checkboxRow}>
                    <input
                      type="checkbox"
                      checked={renameDepartment}
                      onChange={(event) => setRenameDepartment(event.target.checked)}
                    />
                    <span>Do you want to rename Department?</span>
                  </label>

                  {renameDepartment && (
                    <div className={styles.fieldGrid}>
                      <label className={styles.field}>
                        <span className={styles.label}>Select Department</span>
                        <select
                          value={selectedDepartment}
                          onChange={(event) => setSelectedDepartment(event.target.value)}
                          className={styles.select}
                          disabled={!selectedBU}
                        >
                          <option value="">Select Department</option>
                          {departmentOptions.map((option) => (
                            <option key={option} value={option}>{option}</option>
                          ))}
                        </select>
                      </label>

                      <label className={styles.field}>
                        <span className={styles.label}>Select target Department</span>
                        <select
                          value={targetDepartment}
                          onChange={(event) => setTargetDepartment(event.target.value)}
                          className={styles.select}
                          disabled={!targetBU}
                        >
                          <option value="">Select target Department</option>
                          {targetDepartmentOptions.map((option) => (
                            <option key={option} value={option}>{option}</option>
                          ))}
                        </select>
                      </label>
                    </div>
                  )}
                </div>
              )}

              {hasCompletedDepartmentStep && (
                <div className={styles.card}>
                  <label className={styles.checkboxRow}>
                    <input
                      type="checkbox"
                      checked={renameSubDepartment}
                      onChange={(event) => setRenameSubDepartment(event.target.checked)}
                      disabled={!subDepartmentFieldName}
                    />
                    <span>Do you want to rename Sub-Department?</span>
                  </label>

                  {!subDepartmentFieldName && (
                    <p className={styles.helperText}>Sub-Department field was not detected in this library, so this option is unavailable.</p>
                  )}

                  {renameSubDepartment && subDepartmentFieldName && (
                    <div className={styles.fieldGrid}>
                      <label className={styles.field}>
                        <span className={styles.label}>Select Sub-Department</span>
                        <select
                          value={selectedSubDepartment}
                          onChange={(event) => setSelectedSubDepartment(event.target.value)}
                          className={styles.select}
                          disabled={!selectedDepartment}
                        >
                          <option value="">Select Sub-Department</option>
                          {subDepartmentOptions.map((option) => (
                            <option key={option} value={option}>{option}</option>
                          ))}
                        </select>
                      </label>

                      <label className={styles.field}>
                        <span className={styles.label}>Select target Sub-Department</span>
                        <select
                          value={targetSubDepartment}
                          onChange={(event) => setTargetSubDepartment(event.target.value)}
                          className={styles.select}
                          disabled={!targetDepartment}
                        >
                          <option value="">Select target Sub-Department</option>
                          {targetSubDepartmentOptions.map((option) => (
                            <option key={option} value={option}>{option}</option>
                          ))}
                        </select>
                      </label>
                    </div>
                  )}
                </div>
              )}

              {confirmationCount !== null && (
                <div className={styles.confirmationCard}>
                  <strong>This will update {confirmationCount} documents.</strong>
                  <div className={styles.confirmationActions}>
                    <button type="button" className={styles.secondaryButton} onClick={() => setConfirmationCount(null)} disabled={submitting}>
                      Cancel
                    </button>
                    <button type="button" className={styles.primaryButton} onClick={handleConfirmUpdate} disabled={submitting}>
                      {submitting ? 'Updating...' : 'Confirm Update'}
                    </button>
                  </div>
                </div>
              )}
            </>
          )}
        </div>

        <div className={styles.footer}>
          <button type="button" className={styles.secondaryButton} onClick={handleClose} disabled={submitting}>
            Close
          </button>
          <button type="button" className={styles.primaryButton} onClick={handlePrepareUpdate} disabled={loading || submitting}>
            Update Changes
          </button>
        </div>
      </div>
  );

  if (variant === 'inline') {
    return <div className={styles.inlineHost}>{dialogContent}</div>;
  }

  return (
    <div className={styles.overlay}>
      {dialogContent}
    </div>
  );
};

export default RenameBusinessUnitDialog;
