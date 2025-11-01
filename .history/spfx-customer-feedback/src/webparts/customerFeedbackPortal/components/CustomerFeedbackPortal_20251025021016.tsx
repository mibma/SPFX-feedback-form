import * as React from 'react';
import styles from './CustomerFeedbackPortal.module.scss';
import type { ICustomerFeedbackPortalProps } from './ICustomerFeedbackPortalProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField, PrimaryButton, Dropdown, IDropdownOption } from '@fluentui/react';

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

  private handleSubmit = async () => {
    const { name, age, email, rate, comments } = this.state;

    if (!name || !email || !age || !rate) {
      alert('Please fill all required fields.');
      return;
    }

    this.setState({ submitting: true });

    try {
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

  public render(): React.ReactElement<ICustomerFeedbackPortalProps> {
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
        <div className={styles.welcome}>
          <h2>Welcome, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>

        <div className={styles.feedbackForm}>
          <h3>Customer Feedback Form</h3>
          <TextField label="Customer Name" required value={name} onChange={(_, v) => this.setState({ name: v || '' })} />
          <TextField label="Age" required type="number" value={age} onChange={(_, v) => this.setState({ age: v || '' })} />
          <TextField label="Email" required value={email} onChange={(_, v) => this.setState({ email: v || '' })} />
          <Dropdown label="Rating (1-5)" selectedKey={rate} options={rateOptions} onChange={(_, option) => this.setState({ rate: option?.key as number })} />
          <TextField label="Comments" multiline value={comments} onChange={(_, v) => this.setState({ comments: v || '' })} />
          <PrimaryButton text={submitting ? "Submitting..." : "Submit Feedback"} onClick={this.handleSubmit} disabled={submitting} />
        </div>
      </section>
    );
  }
}
