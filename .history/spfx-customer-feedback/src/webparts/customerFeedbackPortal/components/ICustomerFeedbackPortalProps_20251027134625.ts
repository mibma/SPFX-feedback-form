// src/webparts/customerFeedbackPortal/components/ICustomerFeedbackPortalProps.ts
import { spfi } from '@pnp/sp';

export interface ICustomerFeedbackPortalProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  sp: ReturnType<typeof spfi>;
}
