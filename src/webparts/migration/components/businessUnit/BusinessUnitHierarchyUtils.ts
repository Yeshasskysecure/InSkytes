import { BUSINESS_UNIT_HIERARCHY } from './RenameBusinessUnitDialog/BusinessUnitHierarchy';

export const NONE_OPTION = 'None';

const normalizeValue = (value?: string): string => (value || '').trim();

const isMeaningfulValue = (value?: string): boolean => {
  const normalized = normalizeValue(value);
  return normalized.length > 0 && normalized !== '-';
};

const toUniqueSortedValues = (values: string[]): string[] =>
  Array.from(new Set(values.map(normalizeValue).filter(isMeaningfulValue))).sort((left, right) => left.localeCompare(right));

export const getBusinessUnitOptions = (): string[] =>
  Object.keys(BUSINESS_UNIT_HIERARCHY).sort((left, right) => left.localeCompare(right));

export const getDepartmentOptions = (businessUnit?: string): string[] => {
  const options = toUniqueSortedValues(Object.keys(BUSINESS_UNIT_HIERARCHY[businessUnit || ''] || {}));
  return options.length > 0 ? options : [NONE_OPTION];
};

export const getSubDepartmentOptions = (businessUnit?: string, department?: string): string[] => {
  const options = toUniqueSortedValues(((BUSINESS_UNIT_HIERARCHY[businessUnit || ''] || {})[department || ''] || []).slice());
  return options.length > 0 ? options : [NONE_OPTION];
};

export const buildBuDepartmentValue = (
  businessUnit?: string,
  department?: string,
  subDepartment?: string
): string => {
  const parts = [businessUnit, department, subDepartment]
    .map(normalizeValue)
    .filter(isMeaningfulValue);

  return parts.join(' / ');
};

export const getBuDepartmentOptions = (): string[] => {
  const options: string[] = [];

  Object.keys(BUSINESS_UNIT_HIERARCHY).forEach((businessUnit) => {
    const normalizedBusinessUnit = normalizeValue(businessUnit);
    if (!isMeaningfulValue(normalizedBusinessUnit)) {
      return;
    }

    options.push(normalizedBusinessUnit);

    const departments = BUSINESS_UNIT_HIERARCHY[businessUnit] || {};
    Object.keys(departments).forEach((department) => {
      const normalizedDepartment = normalizeValue(department);
      if (!isMeaningfulValue(normalizedDepartment)) {
        return;
      }

      options.push(buildBuDepartmentValue(normalizedBusinessUnit, normalizedDepartment));

      (departments[department] || []).forEach((subDepartment) => {
        const normalizedSubDepartment = normalizeValue(subDepartment);
        if (!isMeaningfulValue(normalizedSubDepartment)) {
          return;
        }

        options.push(buildBuDepartmentValue(normalizedBusinessUnit, normalizedDepartment, normalizedSubDepartment));
      });
    });
  });

  return toUniqueSortedValues(options);
};
