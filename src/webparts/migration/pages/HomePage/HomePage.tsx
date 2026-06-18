import * as React from 'react';
import { IHomePageProps } from './IHomePageProps';
import Migration from '../../components/Migration';
import { IMigrationProps } from '../../components/IMigrationProps';

export const HomePage: React.FunctionComponent<IHomePageProps> = (props) => {
  const migrationProps: IMigrationProps = {
    description: props.description || 'Migration',
    context: props.context
  };

  return <Migration {...migrationProps} />;
};



