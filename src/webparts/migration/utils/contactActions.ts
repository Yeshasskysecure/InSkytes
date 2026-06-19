export const IKNOWLEDGE_CONTACT_EMAIL = 'j.yeshas@skysecure.ai';
export const IKNOWLEDGE_CONTACT_SUBJECT = 'Query :';

export const getOutlookComposeUrl = (email: string = IKNOWLEDGE_CONTACT_EMAIL): string =>
  `https://outlook.office.com/mail/deeplink/compose?to=${encodeURIComponent(email)}&subject=${encodeURIComponent(IKNOWLEDGE_CONTACT_SUBJECT)}`;

export const openOutlookCompose = (email: string = IKNOWLEDGE_CONTACT_EMAIL): void => {
  window.open(getOutlookComposeUrl(email), '_blank', 'noopener,noreferrer');
};
