// src/webparts/customerFeedbackPortal/components/CustomerFeedbackPortal.tsx
import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './CustomerFeedbackPortal.module.scss';
import { TextField, PrimaryButton, Rating, RatingSize, Dropdown, IDropdownOption } from '@fluentui/react';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/lists/web";
import { getSP } from "../pnpConfigFile";
import "@pnp/sp/site-users/web";

interface CustomerFeedbackPortalProps {
  listName?: string; // optional, default 'cloudlist'
  description?: string;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string; // name shown in "Welcome, <name>"
}

const CustomerFeedbackPortal: React.FC<CustomerFeedbackPortalProps> = ({
  listName = 'cloudlist',
  description,
  environmentMessage,
  hasTeamsContext,
  userDisplayName
}) => {
  const [name, setName] = useState<string>(userDisplayName ?? '');
  const [email, setEmail] = useState<string>(''); // user must type this now
  const [rate, setRate] = useState<number>(0);
  const [comments, setComments] = useState<string>('');
  const [service, setService] = useState<string>(''); // Service dropdown state
  const [submitting, setSubmitting] = useState<boolean>(false);
  const [message, setMessage] = useState<string | undefined>(undefined);
  const [messageType, setMessageType] = useState<'error' | 'success' | 'info' | undefined>(undefined);

  // Service dropdown options
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
        const currentUser = await sp.web.currentUser();
        if (currentUser && !userDisplayName) {
          setName(currentUser.Title);
        }
      } catch (err: unknown) {
        console.warn('Could not fetch current user:', err);
      }
    };
  
    void fetchCurrentUser(); // fix floating promise
  
    if (userDisplayName) setName(userDisplayName);
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

    const hasField = (internalName: string): boolean =>
      fieldsArray.some(f => f.InternalName === internalName);

    // ✅ Replace `any` with Record<string, unknown>
    const itemData: Record<string, unknown> = {};

    if (hasField('Title')) itemData.Title = name;

    const custNameCandidates = ['CustomerName', 'Customer_x0020_Name', 'Customer_name', 'Customer'];
    for (const n of custNameCandidates) {
      if (hasField(n)) {
        itemData[n] = name;
        break;
      }
    }

    const emailCandidates = ['Email', 'email', 'CustomerEmail', 'Customer_x0020_Email'];
    for (const n of emailCandidates) {
      if (hasField(n)) {
        itemData[n] = email.trim();
        break;
      }
    }

    const rateCandidates = ['Rate', 'rate', 'Rating', 'rating'];
    for (const n of rateCandidates) {
      if (hasField(n)) {
        itemData[n] = rate;
        break;
      }
    }

    const commentsCandidates = ['Comments', 'comments', 'CustomerComments'];
    for (const n of commentsCandidates) {
      if (hasField(n)) {
        itemData[n] = comments;
        break;
      }
    }

    const serviceCandidates = ['Service', 'service', 'Service_x0020_Type'];
    for (const n of serviceCandidates) {
      if (hasField(n)) {
        itemData[n] = service || null;
        break;
      }
    }

    if (Object.keys(itemData).length === 0) itemData.Title = name;

    await list.items.add(itemData);

    setMessage('Thank you for your feedback!');
    setMessageType('success');
    setEmail('');
    setRate(0);
    setComments('');
    setService('');
  } catch (err: unknown) {
    console.error('Submit error:', err);
    let errorMessage = 'Error submitting feedback. ';
    if (err instanceof Error) errorMessage += err.message;
    setMessage(errorMessage);
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

{/* Pre-filled name (not editable) */}
<TextField
  label="Customer Name"
  required
  value={name}
  disabled
/>

{/* Email must be entered by user */}
<TextField
  label="Email"
  required
  value={email}
  onChange={(_e, v) => setEmail(v || '')}
/>

{/* Service dropdown (required) */}
<div style={{ marginTop: 8 }}>
  <label style={{ display: 'block', marginBottom: 4 }}>
    Service <span style={{ color: '#a4262c' }}>*</span>
  </label>
  <Dropdown
    placeholder="Select a service"
    selectedKey={service || undefined}
    options={serviceOptions}
    onChange={(_e, option) => setService((option?.key as string) || '')}
    styles={{ dropdown: { maxWidth: 320 } }}
  />
</div>

<div style={{ marginTop: 12 }}>
  <label style={{ display: 'block', marginBottom: 4 }}>
    Rating <span style={{ color: '#a4262c' }}>*</span>
  </label>
  <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
    <Rating
      rating={rate}
      max={5}
      size={RatingSize.Large}
      allowZeroStars={true}
      onChange={(_e, newValue) => setRate(newValue ?? 0)}
      styles={() => ({
        root: {
          selectors: {
            '.ms-Rating-starFront': { color: '#20B2AA' },
            ':hover .ms-Rating-starFront': { color: '#20B2AA' }
          }
        },
        ratingStarFront: { color: '#20B2AA' },
        ratingButton: {
          selectors: {
            ':hover .ms-Rating-starFront': { color: '#20B2AA' },
            ':focus .ms-Rating-starFront': { color: '#20B2AA' }
          }
        },
        ratingStarBack: { color: '#e1dfdd' }
      })}
    />
    <span style={{ color: '#605e5c', fontSize: 14 }}>
      {rate > 0 ? `${rate}/5` : '-/5'}
    </span>
  </div>
</div>

<TextField
  label="Comments"
  multiline
  value={comments}
  onChange={(_e, v) => setComments(v || '')}
/>

{/* Submit button */}
<div style={{ marginTop: 20, textAlign: 'center' }}>
  <PrimaryButton
    text={submitting ? 'Submitting...' : 'Submit Feedback'}
    onClick={handleSubmit}
    disabled={submitting}
    style={{ minWidth: 150 }}
  />
</div>

{/* ✅ Move success/failure message BELOW the button */}
{message && (
  <div
    className={
      messageType === 'error'
        ? styles.messageError
        : styles.messageSuccess
    }
    style={{
      marginTop: 12,
      textAlign: 'center',
      fontWeight: 500
    }}
  >
    {message}
  </div>
)}

</div>
</section>
  );
};

export default CustomerFeedbackPortal;
