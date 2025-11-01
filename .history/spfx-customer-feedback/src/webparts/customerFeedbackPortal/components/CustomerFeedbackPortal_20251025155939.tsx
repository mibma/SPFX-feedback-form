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
  rate: number | undefined;
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

  // explicit return type to satisfy eslint/ts rules
  private handleSubmit = async (): Promise<void> => {
    const { name, age, email, rate, comments } = this.state;

    if (!name || !email || !age || !rate) {
      alert('Please fill all required fields.');
      return;
    }

    this.setState({ submitting: true });

    try {
      // use the sp instance passed from the web part
      await this.props.sp.web.lists.getByTitle('Feedback-list').items.add({
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
      alert('Error submitting feedback.');
      this.setState({ submitting: false });
    }
  };

  // small typed onChange helpers (not required but clearer)
  private onChangeText = (key: keyof ICustomerFeedbackPortalState) => (_: any, newValue?: string): void => {
    // Age and rate are strings/numbers handled in state; keep type conversions where needed
    this.setState({ [key]: newValue || '' } as Pick<ICustomerFeedbackPortalState, keyof ICustomerFeedbackPortalState>);
  };

  public render(): React.ReactElement<ICustomerFeedbackPortalProps> {
    // keep and use description + environmentMessage so they are not unused
    const { description, isDarkTheme, environmentMessage, hasTeamsContext, userDisplayName } = this.props;
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
          <TextField
            label="Customer Name"
            required
            value={name}
            onChange={this.onChangeText('name')}
          />

          <TextField
            label="Age"
            required
            type="number"
            value={age}
            onChange={this.onChangeText('age')}
          />

          <TextField
            label="Email"
            required
            value={email}
            onChange={this.onChangeText('email')}
          />

          <Dropdown
            label="Rating (1-5)"
            selectedKey={rate}
            options={rateOptions}
            onChange={(_, option) => this.setState({ rate: option?.key as number })}
          />

          <TextField
            label="Comments"
            multiline
            value={comments}
            onChange={(_, v) => this.setState({ comments: v || '' })}
          />

          <PrimaryButton
            text={submitting ? "Submitting..." : "Submit Feedback"}
            onClick={this.handleSubmit}
            disabled={submitting}
          />
        </div>
      </section>
    );
  }
}
