import { SPFI } from "@pnp/sp";
import { ReturnType } from '@microsoft/whatever'; // not needed; below works
import { spfi } from "@pnp/sp";

export interface ICustomerFeedbackPortalProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  sp: SPFI;
}
