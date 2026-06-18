import { PRODUCTION_GUIDS } from '../config/productionConfig';

export const SIDEBAR_GROUP_IDS = {
  admins: PRODUCTION_GUIDS.roleGroups.admins,
  approvers: PRODUCTION_GUIDS.roleGroups.approvers,
  contributors: PRODUCTION_GUIDS.roleGroups.contributors,
  learners: PRODUCTION_GUIDS.roleGroups.learners
} as const;
