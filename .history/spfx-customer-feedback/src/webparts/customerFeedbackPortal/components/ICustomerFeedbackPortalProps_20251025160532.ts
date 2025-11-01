// src/webparts/customerFeedbackPortal/components/ICustomerFeedbackPortalProps.ts
import { spfi } from '@pnp/sp';

/**
 * Props passed from the web part into the React component.
 * Make sure this matches what the web part passes (description, isDarkTheme, etc).
 */
export interface ICustomerFeedbackPortalProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  sp: any;
}
