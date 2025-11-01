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

  private handleSubmit = async (): Promise<void> => {
    const { name, age, email, rate, comments } = this.state;

    if (!name || !email || !age || !rate) {
      alert('Please fill all required fields.');
      return;
    }

    this.setState({ submitting: true });

    try {
      await this.props.sp.web.lists.getByTitle('Feedback Form').items.add({
        Customer_name: name,
        Age: Number(age),
        Email: email,
        Rate: rate,
        Comments: comments
      });

      alert('Thank you for your feedback!');

      // Reset form
      this.setState({
        name: '',
        age: '',
        email: '',
        rate: 5,
        comments: '',
        submitting: false
      });

    } catch (err) {
      console.error(err);
      alert('Error submitting feedback. See console for details.');
      this.setState({ submitting: false });
    }
  };

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
