/**
 * Reusable action card: Title, description, and a primary action button.
 */

import type { ReactNode } from 'react';
import {
  Body1,
  Button,
  Card,
  CardFooter,
  CardHeader,
  Title3,
} from '@fluentui/react-components';

interface ActionCardProps {
  title: string;
  description: string;
  buttonLabel: string;
  onAction: () => void;
  /** Optional extra content rendered after the CardFooter (e.g. VerifyResultPanel). */
  children?: ReactNode;
}

export function ActionCard({
  title,
  description,
  buttonLabel,
  onAction,
  children,
}: ActionCardProps) {
  return (
    <Card>
      <CardHeader
        header={<Title3>{title}</Title3>}
        description={<Body1>{description}</Body1>}
      />
      <CardFooter
        action={
          <Button appearance="primary" onClick={onAction}>
            {buttonLabel}
          </Button>
        }
      />
      {children}
    </Card>
  );
}
