// src/webparts/customerFeedbackPortal/components/CustomerFeedbackPortal.tsx
import * as React from 'react';
import { useState } from 'react';
import styles from './CustomerFeedbackPortal.module.scss';
import { TextField, PrimaryButton, Rating, RatingSize } from '@fluentui/react';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/lists/web";
import { getSP } from "../pnpConfigFile";

interface CustomerFeedbackPortalProps {
  listName?: string; // optional, default 'cloudlist'
  description?: string;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string;
}

const CustomerFeedbackPortal: React.FC<CustomerFeedbackPortalProps> = ({
  listName = 'cloudlist',
  description,
  environmentMessage,
  hasTeamsContext,
  userDisplayName
}) => {
  const [name, setName] = useState<string>('');
  const [age, setAge] = useState<string>('');
  const [email, setEmail] = useState<string>('');
  const [rate, setRate] = useState<number>(0);
  const [comments, setComments] = useState<string>('');
  const [submitting, setSubmitting] = useState<boolean>(false);
  const [message, setMessage] = useState<string | undefined>(undefined);
  const [messageType, setMessageType] = useState<'error' | 'success' | 'info' | undefined>(undefined);

  // Using Fluent UI Rating control instead of a dropdown

  const isValidEmail = (e: string): boolean => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);

  const handleSubmit = async (): Promise<void> => {
    setMessage(undefined);
    setMessageType(undefined);

    // basic validation
    if (!name.trim() || !age.trim() || !email.trim()) {
      setMessage('Please fill all required fields (Name, Age, Email).');
      setMessageType('error');
      return;
    }
    if (!isValidEmail(email.trim())) {
      setMessage('Please enter a valid email address.');
      setMessageType('error');
      return;
    }
    const ageNum = Number(age);
    if (isNaN(ageNum) || ageNum <= 0 || ageNum > 120) {
      setMessage('Please enter a realistic age.');
      setMessageType('error');
      return;
    }
    if (rate < 1 || rate > 5) {
      setMessage('Rating must be between 1 and 5.');
      setMessageType('error');
      return;
    }

    setSubmitting(true);

    try {
      const sp = getSP();

      // verify list exists & fetch fields
      const list = sp.web.lists.getByTitle(listName);
      const listInfo = await list.select('Title')();
      console.log('List found:', listInfo?.Title);

      const fieldsArray: Array<{ InternalName: string; Title?: string }> =
        await sp.web.lists.getByTitle(listName).fields.select('InternalName', 'Title')();

      console.log('Available fields:', fieldsArray);

      const hasField = (internalName: string): boolean => {
        for (let i = 0; i < fieldsArray.length; i++) {
          if (fieldsArray[i].InternalName === internalName) {
            return true;
          }
        }
        return false;
      };

      // Build payload using likely internal names; also include Title by default if available
      const itemData: any = {};

      // Title: many lists require Title; set to name (optional if your list doesn't have Title)
      if (hasField('Title')) {
        itemData.Title = name;
      }

      // CUSTOMER NAME - try multiple guessed internal names
      const custNameCandidates = [
        'CustomerName', 'Customer_x0020_Name', 'Customer_x0020_name', 'Customer_name',
        'Customer', 'customername', 'Customer_x0020_name0'
      ];
      for (const n of custNameCandidates) {
        if (hasField(n)) {
          itemData[n] = name;
          break;
        }
      }

      // AGE
      const ageCandidates = ['Age', 'age', 'CustomerAge', 'Customer_x0020_Age'];
      for (const n of ageCandidates) {
        if (hasField(n)) {
          itemData[n] = ageNum;
          break;
        }
      }

      // EMAIL
      const emailCandidates = ['Email', 'email', 'CustomerEmail', 'Customer_x0020_Email'];
      for (const n of emailCandidates) {
        if (hasField(n)) {
          itemData[n] = email.trim();
          break;
        }
      }

      // RATE
      const rateCandidates = ['Rate', 'rate', 'Rating', 'rating', 'CustomerRate', 'Customer_x0020_Rate'];
      for (const n of rateCandidates) {
        if (hasField(n)) {
          itemData[n] = rate;
          break;
        }
      }

      // COMMENTS
      const commentsCandidates = ['Comments', 'comments', 'CustomerComments', 'Customer_x0020_Comments'];
      for (const n of commentsCandidates) {
        if (hasField(n)) {
          itemData[n] = comments;
          break;
        }
      }

      // If none of the candidate fields exist except Title, ensure there's at least something to add.
      if (Object.keys(itemData).length === 0) {
        // fallback: try Title only
        itemData.Title = name;
      }

      console.log('Adding item with payload:', itemData);

      const addResult = await list.items.add(itemData);
      console.log('Add result:', addResult);

      setMessage('Thank you for your feedback!');
      setMessageType('success');

      // reset fields
      setName('');
      setAge('');
      setEmail('');
      setRate(0);
      setComments('');
    } catch (err: any) {
      console.error('Submit error:', err);
      let errorMessage = 'Error submitting feedback. ';
      if (err?.status === 404) {
        errorMessage += `List "${listName}" not found.`;
      } else if (err?.status === 403) {
        errorMessage += 'Permission denied. Check your list permissions.';
      } else if (err?.message) {
        errorMessage += err.message;
      } else {
        errorMessage += 'See console for details.';
      }
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
          <div>Welcome, <strong>{userDisplayName ?? ''}</strong></div>
          {environmentMessage && <div className={styles.env}>{environmentMessage}</div>}
          {description && <div className={styles.description}>Web part property value: <strong>{description}</strong></div>}
        </div>
      </div>

      <div className={styles.feedbackForm}>
        {message && (
          <div className={messageType === 'error' ? styles.messageError : styles.messageSuccess}>
            {message}
          </div>
        )}

        <TextField label="Customer Name" required value={name} onChange={(_e, v) => setName(v || '')} />
        <TextField label="Age" required value={age} onChange={(_e, v) => setAge(v || '')} />
        <TextField label="Email" required value={email} onChange={(_e, v) => setEmail(v || '')} />
        <div>
          <label style={{ display: 'block', marginBottom: 4 }}>Rating (1-5)</label>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <Rating
              rating={rate}
              max={5}
              size={RatingSize.Large}
              allowZeroStars={true}
              onChange={(_e, newValue) => setRate(newValue ?? 0)}
              styles={() => ({
                ratingStarFront: {
                  color: '#20B2AA'
                },
                ratingButton: {
                  selectors: {
                    ':hover .ms-Rating-starFront': {
                      color: '#6fd3cd'
                    },
                  
                  }
                },
                ratingStarBack: {
                  color: '#e1dfdd'
                }
              })}
            />
            <span style={{ color: '#605e5c', fontSize: 14 }}>
              {rate > 0 ? `${rate}/5` : '-/5'}
            </span>
          </div>
        </div>
        <TextField label="Comments" multiline value={comments} onChange={(_e, v) => setComments(v || '')} />

        <div style={{ marginTop: 20, textAlign: 'center' }}>
          <PrimaryButton
            text={submitting ? 'Submitting...' : 'Submit Feedback'}
            onClick={handleSubmit}
            disabled={submitting}
            style={{ minWidth: 150 }}
          />
        </div>
      </div>
    </section>
  );
};

export default CustomerFeedbackPortal;
