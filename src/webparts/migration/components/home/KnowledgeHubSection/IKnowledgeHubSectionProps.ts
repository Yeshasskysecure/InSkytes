import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IKnowledgeHubSectionProps {
    context: WebPartContext;
    showPublishedSection?: boolean;
    isLearner?: boolean;
}
