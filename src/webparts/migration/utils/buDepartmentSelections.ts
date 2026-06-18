import { COLUMN_NAMES } from '../config/appConfig';

export const KM_DATA_HUB_BU_FIELD_INTERNAL_NAME = COLUMN_NAMES.businessUnit;
export const KM_DATA_HUB_DEPARTMENT_FIELD_INTERNAL_NAME = COLUMN_NAMES.department;

export interface ISelectedBuDepartmentTermLike {
  id?: string;
  name?: string;
  children?: ISelectedBuDepartmentTermLike[];
}

const buildTermFieldValue = (terms: Array<ISelectedBuDepartmentTermLike | null | undefined>): string =>
  (terms || [])
    .filter((term): term is ISelectedBuDepartmentTermLike => !!term?.id && !!term?.name)
    .map((term) => `${term.name}|${term.id}`)
    .join(';');

export const getLeafSelectedDepartments = <T extends ISelectedBuDepartmentTermLike>(selectedDepts: T[] = []): T[] => {
  const selectedDeptIds = new Set(selectedDepts.map((term) => term.id).filter(Boolean));

  return selectedDepts.filter((term) => {
    const hasSelectedChild = (term.children || []).some((child) => !!child.id && selectedDeptIds.has(child.id));
    return !hasSelectedChild;
  });
};

export const buildBuDepartmentFieldUpdates = (
  selectedBUs: ISelectedBuDepartmentTermLike[] = [],
  selectedDepts: ISelectedBuDepartmentTermLike[] = [],
  options?: { includeEmpty?: boolean }
): Array<{ FieldName: string; FieldValue: string }> => {
  const leafDepts = getLeafSelectedDepartments(selectedDepts);
  const includeEmpty = options?.includeEmpty === true;
  const fieldUpdates = [
    {
      FieldName: KM_DATA_HUB_BU_FIELD_INTERNAL_NAME,
      FieldValue: buildTermFieldValue(selectedBUs)
    },
    {
      FieldName: KM_DATA_HUB_DEPARTMENT_FIELD_INTERNAL_NAME,
      FieldValue: buildTermFieldValue(leafDepts)
    }
  ];

  return includeEmpty ? fieldUpdates : fieldUpdates.filter((field) => !!field.FieldValue);
};
