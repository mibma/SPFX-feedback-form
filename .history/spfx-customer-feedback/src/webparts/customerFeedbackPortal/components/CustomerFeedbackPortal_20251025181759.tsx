// src/webparts/customerFeedbackPortal/components/CustomerFeedbackPortal.tsx
import * as React from 'react';
import styles from './CustomerFeedbackPortal.module.scss';
import type { ICustomerFeedbackPortalProps } from './ICustomerFeedbackPortalProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField, PrimaryButton, Dropdown, IDropdownOption } from '@fluentui/react';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

interface ICustomerFeedbackPortalState {
  name: string;
  age: string;
  email: string;
  rate: number;
  comments: string;
  submitting: boolean;
}

export default class CustomerFeedbackPortal extends React.Component<ICustomerFeedbackPortalProps, ICustomerFeedbackPortalState> {
  constructor(props: ICustomerFeedbackPortalProps) {
    super(props);

    this.state = {
      name: '',
      age: '',
      email: '',
      rate: 5,
      comments: '',
      submitting: false
    };
  }

  // Explicit handlers for each field (type-safe)
  private onNameChange = (_: any, newValue?: string): void => {
    this.setState({ name: newValue || '' });
  };

  private onAgeChange = (_: any, newValue?: string): void => {
    this.setState({ age: newValue || '' });
  };

  private onEmailChange = (_: any, newValue?: string): void => {
    this.setState({ email: newValue || '' });
  };

  private onCommentsChange = (_: any, newValue?: string): void => {
    this.setState({ comments: newValue || '' });
  };

  private onRateChange = (_: any, option?: IDropdownOption): void => {
    this.setState({ rate: (option?.key as number) || 1 });
  };

  private isValidEmail(email: string): boolean {
    const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return re.test(email);
  }
  
  private async ensureListExists(listTitle: string): Promise<boolean> {
    try {
      const lists = await this.props.sp.web.lists.select("Title").filter(`Title eq '${listTitle}'`).get();
      return lists.length > 0;
    } catch (err) {
      console.error('Error checking list existence', err);
      return false;
    }
  }
  
  private handleSubmit = async (): Promise<void> => {
    const { name, age, email, rate, comments } = this.state;
  
    // client-side validation
    if (!name || !email || !age || !rate) {
      this.showMessage('Please fill all required fields.', 'error');
      return;
    }
    if (!this.isValidEmail(email)) {
      this.showMessage('Please enter a valid email address.', 'error');
      return;
    }
    if (Number(age) <= 0 || Number(age) > 120) {
      this.showMessage('Please enter a realistic age.', 'error');
      return;
    }
    if (rate < 1 || rate > 5) {
      this.showMessage('Rating must be between 1 and 5.', 'error');
      return;
    }
  
    this.setState({ submitting: true });
  
    // optional: check list exists
    const listOk = await this.ensureListExists('Feedback-list');
    if (!listOk) {
      this.showMessage('Feedback Form not found on this site. Please create it or contact admin.', 'error');
      this.setState({ submitting: false });
      return;
    }
  
    try {
      await this.props.sp.web.lists.getByTitle('Feedbacklist').items.add({
        Customer_Name: name,
        Age: Number(age),
        Email: email,
        Rate: rate,
        Comments: comments
      });
      this.showMessage('Thank you — your feedback was submitted.', 'success');
  
      // reset
      this.setState({ name:'', age:'', email:'', rate:5, comments:'', submitting:false });
    } catch (err) {
      console.error('Submit error', err);
      this.showMessage('Error submitting feedback. See console for details.', 'error');
      this.setState({ submitting: false });
    }
  };
  private showMessage(message: string, type: 'error'|'success'|'info') {
    // implement with Fluent UI MessageBar component or a simple alert
    alert(message); // replace with MessageBar for better UX
  }
    
  public render(): React.ReactElement<ICustomerFeedbackPortalProps> {
    const { description, environmentMessage, hasTeamsContext, userDisplayName } = this.props;
    const { name, age, email, rate, comments, submitting } = this.state;

    const rateOptions: IDropdownOption[] = [
      { key: 1, text: '1' },
      { key: 2, text: '2' },
      { key: 3, text: '3' },
      { key: 4, text: '4' },
      { key: 5, text: '5' }
    ];

    return (
      <section className={`${styles.customerFeedbackPortal} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.header}>
          <h2>Customer Feedback Form</h2>
          <div className={styles.meta}>
            <div>Welcome, <strong>{escape(userDisplayName)}</strong></div>
            {environmentMessage && <div className={styles.env}>{escape(environmentMessage)}</div>}
            <div className={styles.description}>Web part property value: <strong>{escape(description)}</strong></div>
          </div>
        </div>

        <div className={styles.feedbackForm}>
          <TextField label="Customer Name" required value={name} onChange={this.onNameChange} />
          <TextField label="Age" required value={age} onChange={this.onAgeChange} />
          <TextField label="Email" required value={email} onChange={this.onEmailChange} />
          <Dropdown label="Rating (1-5)" selectedKey={rate} options={rateOptions} onChange={this.onRateChange} />
          <TextField label="Comments" multiline value={comments} onChange={this.onCommentsChange} />
          <PrimaryButton text={submitting ? "Submitting..." : "Submit Feedback"} onClick={this.handleSubmit} disabled={submitting} />
        </div>
      </section>
    );
  }
}
