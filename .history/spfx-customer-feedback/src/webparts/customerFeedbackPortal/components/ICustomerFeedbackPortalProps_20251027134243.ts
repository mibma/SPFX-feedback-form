import { Web } from "@pnp/sp/webs";
import { SPFI } from "@pnp/sp";

export interface ICustomerFeedbackPortalProps {
  sp: SPFI; // Initialized PnP SP object
  description: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  environmentMessage: string;
}
