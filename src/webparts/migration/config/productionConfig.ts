import { GROUP_IDS, LIST_GUIDS, TERM_SET_IDS } from './appConfig';

export const PRODUCTION_GUIDS = {
  roleGroups: {
    admins: GROUP_IDS.admins,
    approvers: GROUP_IDS.approvers,
    contributors: GROUP_IDS.contributors,
    learners: GROUP_IDS.learners
  },
  userGroups: {
    kmsUsers: GROUP_IDS.kmsUsers
  },
  lists: {
    kmDataHub: LIST_GUIDS.kmDataHub
  },
  termSets: {
    buDepartment: TERM_SET_IDS.buDepartment,
    client: TERM_SET_IDS.client,
    diseaseArea: TERM_SET_IDS.diseaseArea,
    documentType: TERM_SET_IDS.documentType,
    geography: TERM_SET_IDS.geography,
    therapyArea: TERM_SET_IDS.therapyArea
  }
} as const;

