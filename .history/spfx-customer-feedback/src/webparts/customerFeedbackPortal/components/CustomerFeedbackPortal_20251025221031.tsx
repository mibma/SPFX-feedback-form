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
  message?: string;
  messageType?: 'error' | 'success' | 'info';
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
      submitting: false,
      message: undefined,
      messageType: undefined
    };
  }

  // TextField onChange signature: (event?: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => void
  private onNameChange = (_ev?: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ name: newValue || '' });
  };

  private onAgeChange = (_ev?: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ age: newValue || '' });
  };

  private onEmailChange = (_ev?: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ email: newValue || '' });
  };

  private onCommentsChange = (_ev?: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ comments: newValue || '' });
  };

  // Dropdown onChange signature: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => void
  private onRateChange = (_ev: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    this.setState({ rate: (option?.key as number) ?? 1 });
  };

  // explicit return type
  private isValidEmail(email: string): boolean {
    const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return re.test(email);
  }


  // explicit return type
  private async handleSubmit(): Promise<void> {
    const { name, age, email, rate, comments } = this.state;

    // client-side validation
    if (!name || !email || !age || !rate) {
      this.setState({ message: 'Please fill all required fields.', messageType: 'error' });
      return;
    }
    if (!this.isValidEmail(email)) {
      this.setState({ message: 'Please enter a valid email address.', messageType: 'error' });
      return;
    }
    const ageNum = Number(age);
    if (isNaN(ageNum) || ageNum <= 0 || ageNum > 120) {
      this.setState({ message: 'Please enter a realistic age.', messageType: 'error' });
      return;
    }
    if (rate < 1 || rate > 5) {
      this.setState({ message: 'Rating must be between 1 and 5.', messageType: 'error' });
      return;
    }

    this.setState({ submitting: true, message: undefined, messageType: undefined });

    // Connect to the SharePoint list
    const listTitle = 'Feedbacklist';
    
    try {
      // For now, we'll assume the list exists and try to add the item
      // In production, you should deploy this web part to the target site
      await this.props.sp.web.lists.getByTitle(listTitle).items.add({
        Customer_name: name,
        Age: ageNum,
        Email: email,
        Rate: rate,
        Comments: comments
      });

      this.setState({
        name: '',
        age: '',
        email: '',
        rate: 5,
        comments: '',
        submitting: false,
        message: 'Thank you for your feedback!',
        messageType: 'success'
      });
    } catch (err) {
      console.error('Submit error', err);
      this.setState({ submitting: false, message: 'Error submitting feedback. See console.', messageType: 'error' });
    }
  }


  public render(): React.ReactElement<ICustomerFeedbackPortalProps> {
    const { description, environmentMessage, hasTeamsContext, userDisplayName } = this.props;
    const { name, age, email, rate, comments, submitting, message, messageType } = this.state;

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
          {message && (
            <div className={messageType === 'error' ? styles.messageError : styles.messageSuccess}>
              {message}
            </div>
          )}

          <TextField label="Customer Name" required value={name} onChange={this.onNameChange} />
          <TextField label="Age" required value={age} onChange={this.onAgeChange} />
          <TextField label="Email" required value={email} onChange={this.onEmailChange} />
          <Dropdown label="Rating (1-5)" selectedKey={rate} options={rateOptions} onChange={this.onRateChange} />
          <TextField label="Comments" multiline value={comments} onChange={this.onCommentsChange} />

          <div style={{ marginTop: 20, textAlign: 'center' }}>
            <PrimaryButton
              text={submitting ? "Submitting..." : "Submit Feedback"}
              onClick={this.handleSubmit}
              disabled={submitting}
              style={{ minWidth: 150 }}
            />
          </div>

        </div>
      </section>
    );
  }
}
