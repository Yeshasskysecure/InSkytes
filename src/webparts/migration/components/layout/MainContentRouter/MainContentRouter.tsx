import * as React from 'react';
import { IMainContentRouterProps } from './IMainContentRouterProps';
import { AboutPage } from '../../../pages/AboutPage/AboutPage';

export const MainContentRouter: React.FunctionComponent<IMainContentRouterProps> = (props) => {
  return (
    <AboutPage
      context={props.context}
      onSearchOpen={props.onSearchOpen}
      onUploadOpen={props.onUploadOpen}
      onCategorySelect={props.onCategorySelect}
      hideUploadButton={props.hideUploadButton}
      hideCategorySection={props.hideCategorySection}
      isAdmin={props.isAdmin}
      configService={props.configService || undefined}
    />
  );
};
