import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { TERM_SET_IDS } from '../config/appConfig';

export interface Term {
  id: string;
  name: string;
  description?: string;
  children: Term[];
}

interface IRawTermLabel {
  name: string;
}

interface IRawParentRef {
  id?: string;
}

interface IRawTerm {
  id: string;
  labels?: IRawTermLabel[];
  parent?: IRawParentRef;
}

interface ITermStoreResponse {
  value?: IRawTerm[];
}

const BU_DEPARTMENT_TERM_SET_ID = TERM_SET_IDS.buDepartment;
const DOCUMENT_TYPE_TERM_SET_ID = TERM_SET_IDS.documentType;

export const fetchBUDepartmentTerms = async (
  spHttpClient: SPHttpClient,
  siteUrl: string
): Promise<Term[]> => {
  const termSetId = BU_DEPARTMENT_TERM_SET_ID;

  const buResponse: SPHttpClientResponse = await spHttpClient.get(
    `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/children`,
    SPHttpClient.configurations.v1
  );
  const buData = await buResponse.json();
  const buTerms: any[] = buData.value || [];

  const roots: Term[] = await Promise.all(
    buTerms.map(async (bu: any) => {
      const deptResponse: SPHttpClientResponse = await spHttpClient.get(
        `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/terms/${bu.id}/children`,
        SPHttpClient.configurations.v1
      );
      const deptData = await deptResponse.json();
      const deptTerms: any[] = deptData.value || [];

      const departments: Term[] = await Promise.all(
        deptTerms.map(async (dept: any) => {
          const subResponse: SPHttpClientResponse = await spHttpClient.get(
            `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/terms/${dept.id}/children`,
            SPHttpClient.configurations.v1
          );
          const subData = await subResponse.json();
          const subTerms: any[] = subData.value || [];

          return {
            id: dept.id,
            name: dept.labels?.[0]?.name ?? dept.id,
            children: subTerms.map((sub: any) => ({
              id: sub.id,
              name: sub.labels?.[0]?.name ?? sub.id,
              children: []
            }))
          };
        })
      );

      return {
        id: bu.id,
        name: bu.labels?.[0]?.name ?? bu.id,
        children: departments
      };
    })
  );

  console.log('BU roots:', roots.map((r) => r.name));
  console.log('Sample departments under first BU:', roots[0]?.children?.map((d) => d.name));

  return roots;
};

export async function fetchDocumentTypeTerms(
  spHttpClient: any,
  siteUrl: string
): Promise<Term[]> {
  const termSetId = DOCUMENT_TYPE_TERM_SET_ID;

  const response = await spHttpClient.get(
    `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/children?$select=id,labels,descriptions`,
    SPHttpClient.configurations.v1
  );
  const data = await response.json();
  const level1Terms: any[] = data.value || [];

  const result: Term[] = await Promise.all(
    level1Terms.map(async (l1: any) => {
      const l2Response = await spHttpClient.get(
        `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/terms/${l1.id}/children?$select=id,labels,descriptions`,
        SPHttpClient.configurations.v1
      );
      const l2Data = await l2Response.json();
      const l2Terms: any[] = l2Data.value || [];

      const children: Term[] = await Promise.all(
        l2Terms.map(async (l2: any) => {
          const l3Response = await spHttpClient.get(
            `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/terms/${l2.id}/children?$select=id,labels,descriptions`,
            SPHttpClient.configurations.v1
          );
          const l3Data = await l3Response.json();
          const l3Terms: any[] = l3Data.value || [];

          return {
            id: l2.id,
            name: l2.labels?.[0]?.name ?? l2.id,
            description: l2.descriptions?.[0]?.description ?? '',
            children: l3Terms.map((l3: any) => ({
              id: l3.id,
              name: l3.labels?.[0]?.name ?? l3.id,
              description: l3.descriptions?.[0]?.description ?? '',
              children: []
            }))
          };
        })
      );

      return {
        id: l1.id,
        name: l1.labels?.[0]?.name ?? l1.id,
        description: l1.descriptions?.[0]?.description ?? '',
        children
      };
    })
  );

  console.log('Document type terms fetched:', result.length);
  return result;
}
