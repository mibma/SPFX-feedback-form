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

  // On mount: attempt to fetch current user (Title) with PnP but DO NOT auto-fill email.
  useEffect(() => {
    const fetchCurrentUser = async () => {
      try {
        const sp = getSP();
        const currentUser = await (sp as any).web.currentUser.get();
        if (currentUser) {
          // only set name here — do NOT set email automatically
          if (!userDisplayName && currentUser.Title) {
            setName(currentUser.Title);
          }
        }
      } catch (err) {
        // log for debugging — don't crash the component
        // eslint-disable-next-line no-console
        console.warn('Could not fetch current user:', err);
      }
    };

    fetchCurrentUser();

    // keep the displayed name in sync when the prop is provided/changes
    if (userDisplayName) {
      setName(userDisplayName);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [userDisplayName]);

  const isValidEmail = (e: string): boolean => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);

  const handleSubmit = async (): Promise<void> => {
    setMessage(undefined);
    setMessageType(undefined);

    // basic validation: Name (prefilled), Email (user-entered), Service, Rating required
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

      // Title: many lists require Title; set to name
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

      // EMAIL (now from user input)
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

      // SERVICE (choice field)
      const serviceCandidates = [
        'Service', 'service', 'Service_x0020_Type', 'Service_x0020_Name', 'ServiceChoice', 'Service_x0020_Choice'
      ];
      for (const n of serviceCandidates) {
        if (hasField(n)) {
          itemData[n] = service || null;
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

      // reset fields — clear email (user must re-enter next time), other fields cleared too
      setEmail('');
      setRate(0);
      setComments('');
      setService('');
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
          <label style={{ display: 'block', marginBottom: 4 }}>Service <span style={{ color: '#a4262c' }}>*</span></label>
          <Dropdown
            placeholder="Select a service"
            selectedKey={service || undefined}
            options={serviceOptions}
            onChange={(_e, option) => setService((option?.key as string) || '')}
            styles={{ dropdown: { maxWidth: 320 } }}
          />
        </div>

        <div style={{ marginTop: 12 }}>
          <label style={{ display: 'block', marginBottom: 4 }}>Rating (1-5) <span style={{ color: '#a4262c' }}>*</span></label>
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
                    '.ms-Rating-starFront': {
                      color: '#20B2AA'
                    },
                    ':hover .ms-Rating-starFront': {
                      color: '#20B2AA'
                    }
                  }
                },
                ratingStarFront: {
                  color: '#20B2AA'
                },
                ratingButton: {
                  selectors: {
                    ':hover .ms-Rating-starFront': {
                      color: '#20B2AA'
                    },
                    ':focus .ms-Rating-starFront': {
                      color: '#20B2AA'
                    }
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
