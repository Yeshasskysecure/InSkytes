import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from '../../pages/DocumentDetailPage/DocumentDetailPage.module.scss';
import { subscribeToDocumentDataChanged } from '../../services/documentChangeEvents';
import { fetchKMDataHubReadFieldMap } from '../../utils/documentMetadata';
import { BUSINESS_UNIT_HIERARCHY } from '../businessUnit/RenameBusinessUnitDialog/BusinessUnitHierarchy';
import {
  KM_DATA_HUB_BU_FIELD_INTERNAL_NAME,
  KM_DATA_HUB_DEPARTMENT_FIELD_INTERNAL_NAME
} from '../../utils/buDepartmentSelections';
import { LIBRARY_NAMES } from '../../config/appConfig';

export interface IBUDeptDisplayProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  itemId: number;
}

interface IDisplayState {
  businessUnits: string[];
  departments: string[];
}

interface IRenderedDepartmentGroup {
  name: string;
  subDepartments: string[];
}

interface IRenderedBuGroup {
  name: string;
  departments: IRenderedDepartmentGroup[];
}

const LIBRARY_NAME = LIBRARY_NAMES.kmDataHub;

const splitFallbackLabels = (value: string): string[] =>
  value
    .split(/[;,#\r\n]+/)
    .map((entry) => entry.split('|')[0].trim())
    .filter(Boolean);

const extractLabels = (value: any): string[] => {
  if (!value) {
    return [];
  }

  if (typeof value === 'string') {
    return splitFallbackLabels(value);
  }

  if (Array.isArray(value)) {
    return value
      .map((entry) => {
        if (typeof entry === 'string') {
          return entry.trim();
        }

        return (entry?.Label || entry?.label || entry?.name || '').trim();
      })
      .filter(Boolean);
  }

  if (Array.isArray(value?.results)) {
    return extractLabels(value.results);
  }

  return [];
};

const uniqueValues = (values: string[]): string[] =>
  values.filter((value, index, array) => value.length > 0 && array.indexOf(value) === index);

const normalizeComparisonValue = (value: string): string => value.trim().toLowerCase();

const parseDepartmentPath = (
  rawDepartmentValue: string,
  businessUnit: string
): { departmentName: string; subDepartmentName?: string } | null => {
  const segments = rawDepartmentValue
    .split(':')
    .map((segment) => segment.trim())
    .filter(Boolean);

  if (!segments.length) {
    return null;
  }

  const normalizedBusinessUnit = normalizeComparisonValue(businessUnit);
  const normalizedFirstSegment = normalizeComparisonValue(segments[0]);
  const pathSegments =
    normalizedFirstSegment === normalizedBusinessUnit && segments.length > 1
      ? segments.slice(1)
      : segments;

  if (!pathSegments.length) {
    return null;
  }

  if (pathSegments.length === 1) {
    return {
      departmentName: pathSegments[0]
    };
  }

  return {
    departmentName: pathSegments[0],
    subDepartmentName: pathSegments.slice(1).join(': ')
  };
};

const buildRenderedHierarchy = (businessUnits: string[], departments: string[]): IRenderedBuGroup[] => {
  const normalizedDepartments = uniqueValues(departments);
  const unmatchedDepartments = new Set(normalizedDepartments);

  const renderedGroups = businessUnits.map((businessUnit) => {
    const departmentMap = new Map<string, Set<string>>();
    const businessUnitDepartments = BUSINESS_UNIT_HIERARCHY[businessUnit] || {};

    normalizedDepartments.forEach((selectedDepartment) => {
      const parsedDepartmentPath = parseDepartmentPath(selectedDepartment, businessUnit);
      if (!parsedDepartmentPath) {
        return;
      }

      const matchedDepartmentName = Object.keys(businessUnitDepartments).find(
        (departmentName) =>
          normalizeComparisonValue(departmentName) === normalizeComparisonValue(parsedDepartmentPath.departmentName)
      );

      if (!matchedDepartmentName) {
        return;
      }

      if (!departmentMap.has(matchedDepartmentName)) {
        departmentMap.set(matchedDepartmentName, new Set<string>());
      }

      if (parsedDepartmentPath.subDepartmentName) {
        departmentMap.get(matchedDepartmentName)?.add(parsedDepartmentPath.subDepartmentName);
      }

      unmatchedDepartments.delete(selectedDepartment);
    });

    Object.keys(businessUnitDepartments).forEach((departmentName) => {
      const normalizedDepartmentName = normalizeComparisonValue(departmentName);
      const subDepartments = businessUnitDepartments[departmentName] || [];

      normalizedDepartments.forEach((selectedDepartment) => {
        const normalizedSelectedDepartment = normalizeComparisonValue(selectedDepartment);
        if (normalizedSelectedDepartment === normalizedDepartmentName) {
          if (!departmentMap.has(departmentName)) {
            departmentMap.set(departmentName, new Set<string>());
          }
          unmatchedDepartments.delete(selectedDepartment);
          return;
        }

        const matchedSubDepartment = subDepartments.find(
          (subDepartment) => normalizeComparisonValue(subDepartment) === normalizedSelectedDepartment
        );
        if (matchedSubDepartment) {
          if (!departmentMap.has(departmentName)) {
            departmentMap.set(departmentName, new Set<string>());
          }
          departmentMap.get(departmentName)?.add(matchedSubDepartment);
          unmatchedDepartments.delete(selectedDepartment);
        }
      });
    });

    const renderedDepartments = Array.from(departmentMap.entries()).map(([name, subDepartmentSet]) => ({
      name,
      subDepartments: Array.from(subDepartmentSet)
    }));

    return {
      name: businessUnit,
      departments: renderedDepartments
    };
  });

  if (unmatchedDepartments.size > 0) {
    const fallbackDepartments = Array.from(unmatchedDepartments).map((departmentName) => ({
      name: departmentName,
      subDepartments: []
    }));

    if (renderedGroups.length > 0) {
      renderedGroups[0] = {
        ...renderedGroups[0],
        departments: [...renderedGroups[0].departments, ...fallbackDepartments]
      };
    } else {
      renderedGroups.push({
        name: 'Department / Sub Department',
        departments: fallbackDepartments
      });
    }
  }

  return renderedGroups;
};

const renderDepartmentList = (departments: IRenderedDepartmentGroup[]): JSX.Element | null => {
  if (!departments.length) {
    return null;
  }

  return (
    <div className={styles.deptList}>
      {departments.map((department) => (
        <div key={department.name} className={styles.deptEntry}>
          <span className={styles.deptArrow} aria-hidden="true">›</span>
          <div className={styles.deptContent}>
            <span className={styles.deptName}>{department.name}</span>
            {department.subDepartments.length > 0 ? (
              <span className={styles.subDeptList}>
                {department.subDepartments.map((subDepartment) => (
                  <span className={styles.subDeptTag} key={subDepartment}>{subDepartment}</span>
                ))}
              </span>
            ) : null}
          </div>
        </div>
      ))}
    </div>
  );
};

export const BUDeptDisplay: React.FC<IBUDeptDisplayProps> = ({ spHttpClient, siteUrl, itemId }) => {
  const [displayState, setDisplayState] = React.useState<IDisplayState>({
    businessUnits: [],
    departments: []
  });
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState(false);
  const [refreshKey, setRefreshKey] = React.useState(0);

  React.useEffect(() => {
    return subscribeToDocumentDataChanged((event) => {
      if (!event?.documentIds || event.documentIds.indexOf(itemId) !== -1) {
        setRefreshKey((value) => value + 1);
      }
    });
  }, [itemId]);

  React.useEffect(() => {
    let isCancelled = false;

    const loadData = async (): Promise<void> => {
      try {
        setLoading(true);
        setError(false);

        const fieldMap = await fetchKMDataHubReadFieldMap(spHttpClient, siteUrl, LIBRARY_NAME);
        const selectFields = [
          'Id',
          fieldMap.businessUnit,
          fieldMap.department,
          KM_DATA_HUB_BU_FIELD_INTERNAL_NAME,
          KM_DATA_HUB_DEPARTMENT_FIELD_INTERNAL_NAME
        ].filter(Boolean).join(',');

        const [response, textResponse] = await Promise.all([
          spHttpClient.get(
            `${siteUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${itemId})?$select=${selectFields}`,
            SPHttpClient.configurations.v1
          ),
          spHttpClient.get(
            `${siteUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${itemId})/FieldValuesAsText`,
            SPHttpClient.configurations.v1
          )
        ]);
        if (!response.ok) {
          throw new Error(`Failed to fetch BU/Department for item ${itemId}`);
        }

        const rawData = await response.json();
        const data = rawData?.d || rawData || {};
        const textValuesJson = textResponse.ok ? await textResponse.json() : {};
        const textValues = textValuesJson?.d || textValuesJson || {};
        const resolveFieldLabels = (primaryInternalName?: string, fallbackInternalName?: string): string[] => {
          const primaryLabels = uniqueValues(extractLabels(primaryInternalName ? data[primaryInternalName] : null));
          if (primaryLabels.length > 0) {
            return primaryLabels;
          }

          const primaryTextLabels = uniqueValues(extractLabels(primaryInternalName ? textValues[primaryInternalName] : null));
          if (primaryTextLabels.length > 0) {
            return primaryTextLabels;
          }

          const fallbackLabels = uniqueValues(extractLabels(fallbackInternalName ? data[fallbackInternalName] : null));
          if (fallbackLabels.length > 0) {
            return fallbackLabels;
          }

          return uniqueValues(extractLabels(fallbackInternalName ? textValues[fallbackInternalName] : null));
        };
        const businessUnits = resolveFieldLabels(fieldMap.businessUnit, KM_DATA_HUB_BU_FIELD_INTERNAL_NAME);
        const departments = resolveFieldLabels(fieldMap.department, KM_DATA_HUB_DEPARTMENT_FIELD_INTERNAL_NAME);

        if (isCancelled) {
          return;
        }

        setDisplayState({ businessUnits, departments });
        setLoading(false);
      } catch (e) {
        if (isCancelled) {
          return;
        }

        console.error('BUDeptDisplay error:', e);
        setError(true);
        setLoading(false);
      }
    };

    void loadData();

    return () => {
      isCancelled = true;
    };
  }, [itemId, refreshKey, siteUrl, spHttpClient]);

  if (loading) {
    return <div className={styles.loadingState}>Loading BU & Department info...</div>;
  }

  if (error) {
    return <div className={styles.emptyState}>Unable to load BU and Department info</div>;
  }

  if (!displayState.businessUnits.length && !displayState.departments.length) {
    return <div className={styles.emptyState}>No BU or Department assigned</div>;
  }

  const renderedHierarchy = buildRenderedHierarchy(displayState.businessUnits, displayState.departments);

  return (
    <div className={styles.buDeptContainer}>
      {renderedHierarchy.map((businessUnit) => (
        <div key={businessUnit.name} className={styles.buGroup}>
          <div className={styles.buHeader}>
            <span className={styles.buDot} aria-hidden="true" />
            <span className={styles.buName}>{businessUnit.name}</span>
          </div>
          {renderDepartmentList(businessUnit.departments)}
        </div>
      ))}
    </div>
  );
};

export default BUDeptDisplay;
