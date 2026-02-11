/**
 * Load user data from Microsoft Graph API by User Principal Name.
 * Used by the Update flow to populate the UpdateUsers table with current data.
 */

import {
  GRAPH_USERS_URL,
  parseGraphErrorResponse,
  toErrorMessage,
} from './graphHelpers';

const GRAPH_SELECT_FIELDS = [
  'id',
  'mail',
  'employeeId',
  'onPremisesExtensionAttributes',
  'givenName',
  'surname',
  'displayName',
  'country',
  'city',
  'jobTitle',
  'officeLocation',
  'streetAddress',
  'state',
  'postalCode',
  'businessPhones',
  'mobilePhone',
  'companyName',
  'department',
].join(',');

export interface LoadUserResult {
  success: true;
  /** Column values in UpdateUsers table order (Object ID first, then the rest). */
  values: string[];
}

export interface LoadUserError {
  success: false;
  error: string;
}

export type LoadUserOutcome = LoadUserResult | LoadUserError;

interface GraphUserResponse {
  id?: string;
  mail?: string | null;
  employeeId?: string | null;
  onPremisesExtensionAttributes?: {
    extensionAttribute14?: string | null;
    extensionAttribute15?: string | null;
  } | null;
  givenName?: string | null;
  surname?: string | null;
  displayName?: string | null;
  country?: string | null;
  city?: string | null;
  jobTitle?: string | null;
  officeLocation?: string | null;
  streetAddress?: string | null;
  state?: string | null;
  postalCode?: string | null;
  businessPhones?: string[] | null;
  mobilePhone?: string | null;
  companyName?: string | null;
  department?: string | null;
}

/**
 * Loads a user from Graph by UPN and maps the result to UpdateUsers table columns.
 * Returns column values in table order: [ObjectID, UPN, Mail, BMS ID, Local HR ID, SN Ticket ID, ...].
 */
export async function loadUserByUpn(
  upn: string,
  accessToken: string
): Promise<LoadUserOutcome> {
  if (!upn) {
    return { success: false, error: 'User Principal Name is empty' };
  }

  try {
    const url = `${GRAPH_USERS_URL}/${encodeURIComponent(upn)}?$select=${GRAPH_SELECT_FIELDS}`;
    const response = await fetch(url, {
      method: 'GET',
      headers: { Authorization: `Bearer ${accessToken}` },
    });

    if (!response.ok) {
      const errorMessage = await parseGraphErrorResponse(response);
      return { success: false, error: errorMessage };
    }

    const user = (await response.json()) as GraphUserResponse;
    const s = (v: string | null | undefined): string =>
      v != null ? String(v) : '';

    const extAttrs = user.onPremisesExtensionAttributes;

    const values: string[] = [
      s(user.id),                                    // Object ID
      upn,                                           // User Principal Name (keep original)
      s(user.mail),                                  // Mail
      s(user.employeeId),                            // BMS ID
      s(extAttrs?.extensionAttribute14),             // Local HR ID
      s(extAttrs?.extensionAttribute15),             // SN Ticket ID
      s(user.givenName),                             // First Name
      s(user.surname),                               // Last Name
      s(user.displayName),                           // Display Name
      s(user.country),                               // Country
      s(user.city),                                  // City
      s(user.jobTitle),                              // Job Title
      s(user.officeLocation),                        // Office Location
      s(user.streetAddress),                         // Street Address
      s(user.state),                                 // State
      s(user.postalCode),                            // Postal Code
      user.businessPhones?.[0] ?? '',                // Business Phone
      s(user.mobilePhone),                           // Mobile Phone
      s(user.companyName),                           // Company Name
      s(user.department),                            // Department
    ];

    return { success: true, values };
  } catch (err) {
    return { success: false, error: toErrorMessage(err) };
  }
}
