import { SPFI } from "@pnp/sp";


export interface ICustomerFeedbackPortalProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  sp: ReturnType<typeof spfi>; // ensure this line exists
}
