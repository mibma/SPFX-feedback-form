// src/webparts/customerFeedbackPortal/components/CustomerFeedbackPortal.tsx
import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './CustomerFeedbackPortal.module.scss';
import { TextField, PrimaryButton, Rating, RatingSize, Dropdown, IDropdownOption } from '@fluentui/react';
import { getSP } from "../pnpConfigFile";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/lists/web";
import "@pnp/sp/site-users/web";

interface CustomerFeedbackPortalProps {
  listName?: string;
  description?: string;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string;
}

interface IUser {
  Id: number;
  Title: string;
  Email?: string;
}

interface IItemData {
  [key: string]: string | number | null;
}

const CustomerFeedbackPortal: React.FC<CustomerFeedbackPortalProps> = ({
  listName = 'cloudlist',
  description,
  environmentMessage,
  hasTeamsContext,
  userDisplayName
}) => {
  const [name, setName] = useState<string>(userDisplayName ?? '');
  const [email, setEmail] = useState<string>('');
  const [rate, setRate] = useState<number>(0);
  const [comments, setComments] = useState<string>('');
  const [service, setService] = useState<string>('');
  const [submitting, setSubmitting] = useState<boolean>(false);
  const [message, setMessage] = useState<string | undefined>(undefined);
  const [messageType, setMessageType] = useState<'error' | 'success' | 'info' | undefined>(undefined);

  const serviceOptions: IDropdownOption[] = [
    { key: 'Web Hosting', text: 'Web Hosting' },
    { key: 'Network Security', text: 'Network Security' },
    { key: 'Cloud Storage', text: 'Cloud Storage' },
    { key: 'Other', text: 'Other' }
  ];

  useEffect(() => {
    const fetchCurrentUser = async (): Promise<void> => {
      try {
        const sp = getSP();
        const currentUser = await sp.web.currentUser.get();<IUser>();
        if (!userDisplayName && currentUser?.Title) {
          setName(currentUser.Title);
        }
      } catch (err) {
        console.warn('Could not fetch current user:', err);
      }
    };

    void fetchCurrentUser();

    if (userDisplayName) {
      setName(userDisplayName);
    }
  }, [userDisplayName]);

  const isValidEmail = (e: string): boolean => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);

  const handleSubmit = async (): Promise<void> => {
    setMessage(undefined);
    setMessageType(undefined);

    if (!name.trim() || !email.trim() || !service.trim() || rate < 1) {
      setMessage('Please fill all required fields: Name, Email, Service, and Rating (1-5).');
      setMessageType('error');
      return;
    }

    if (!isValidEmail(email.trim())) {
      setMessage('Please enter a valid email address.');
      setMessageType('error');
      return;
    }

    setSubmitting(true);

    try {
      const sp = getSP();
      const list = sp.web.lists.getByTitle(listName);
      const fieldsArray: Array<{ InternalName: string; Title?: string }> =
        await sp.web.lists.getByTitle(listName).fields.select('InternalName', 'Title')();

      const hasField = (internalName: string): boolean => fieldsArray.some(f => f.InternalName === internalName);

      const itemData: IItemData = {};
      if (hasField('Title')) itemData.Title = name;
      if (hasField('CustomerName')) itemData.CustomerName = name;
      if (hasField('Email')) itemData.Email = email;
      if (hasField('Rate')) itemData.Rate = rate;
      if (hasField('Comments')) itemData.Comments = comments;
      if (hasField('Service')) itemData.Service = service;

      await list.items.add(itemData);

      setMessage('Thank you for your feedback!');
      setMessageType('success');

      setEmail('');
      setRate(0);
      setComments('');
      setService('');
    } catch (err: unknown) {
      console.error('Submit error:', err);
      const errorMessage = (err as any)?.message ?? 'See console for details.';
      setMessage(`Error submitting feedback. ${errorMessage}`);
      setMessageType('error');
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <section className={`${styles.customerFeedbackPortal ?? ''} ${hasTeamsContext ? styles.teams ?? '' : ''}`}>
      <div className={styles.header}>
        <h2>Customer Feedback Form</h2>
        <div className={styles.meta}>
          <div>Welcome, <strong>{userDisplayName ?? name ?? ''}</strong></div>
        </div>
      </div>

      <div className={styles.feedbackForm}>
        <TextField label="Customer Name" required value={name} disabled />
        <TextField label="Email" required value={email} onChange={(_e, v) => setEmail(v || '')} />
        <Dropdown
          placeholder="Select a service"
          selectedKey={service || undefined}
          options={serviceOptions}
          onChange={(_e, option) => setService((option?.key as string) || '')}
          styles={{ dropdown: { maxWidth: 320 } }}
        />
        <Rating
          rating={rate}
          max={5}
          size={RatingSize.Large}
          allowZeroStars
          onChange={(_e, newValue) => setRate(newValue ?? 0)}
        />
        <TextField label="Comments" multiline value={comments} onChange={(_e, v) => setComments(v || '')} />

        <div style={{ marginTop: 20, textAlign: 'center' }}>
          <PrimaryButton text={submitting ? 'Submitting...' : 'Submit Feedback'} onClick={handleSubmit} disabled={submitting} />
        </div>

        {message && (
          <div className={messageType === 'error' ? styles.messageError : styles.messageSuccess} style={{ marginTop: 12, textAlign: 'center' }}>
            {message}
          </div>
        )}
      </div>
    </section>
  );
};

export default CustomerFeedbackPortal;
