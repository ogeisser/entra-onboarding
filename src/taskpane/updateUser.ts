/**
 * Update User via Microsoft Graph API (PATCH /users/{objectId}).
 * Maps UpdateUsers table columns to Graph user properties.
 * Empty columns are sent as null to clear the property.
 */

const GRAPH_USERS_URL = 'https://graph.microsoft.com/v1.0/users';

export interface UpdateUserResult {
  status: 'success' | 'error';
  error?: string;
}

/** Column indices matching UPDATE_TABLE_HEADERS (Object ID at index 0). */
const COL = {
  objectId: 0,
  userPrincipalName: 1,
  mail: 2,
  bmsId: 3,
  localHrId: 4,
  snTicketId: 5,
  firstName: 6,
  lastName: 7,
  displayName: 8,
  country: 9,
  city: 10,
  jobTitle: 11,
  officeLocation: 12,
  streetAddress: 13,
  state: 14,
  postalCode: 15,
  businessPhone: 16,
  mobilePhone: 17,
  companyName: 18,
  department: 19,
} as const;

function cell(row: unknown[], index: number): string {
  const raw = row[index];
  return String(raw ?? '').trim();
}

/** Returns string value or null if empty (for Graph PATCH). */
function cellOrNull(row: unknown[], index: number): string | null {
  const v = cell(row, index);
  return v === '' ? null : v;
}

interface GraphUpdateBody {
  displayName: string | null;
  mailNickname?: string;
  userPrincipalName: string | null;
  mail: string | null;
  employeeId: string | null;
  onPremisesExtensionAttributes: {
    extensionAttribute14: string | null;
    extensionAttribute15: string | null;
  };
  givenName: string | null;
  surname: string | null;
  country: string | null;
  city: string | null;
  jobTitle: string | null;
  officeLocation: string | null;
  streetAddress: string | null;
  state: string | null;
  postalCode: string | null;
  businessPhones: string[];
  mobilePhone: string | null;
  companyName: string | null;
  department: string | null;
}

function rowToGraphUpdateBody(row: unknown[]): GraphUpdateBody {
  const upn = cellOrNull(row, COL.userPrincipalName);
  const mailNickname =
    upn && upn.includes('@') ? upn.split('@')[0]! : undefined;

  const businessPhone = cell(row, COL.businessPhone);

  const body: GraphUpdateBody = {
    displayName: cellOrNull(row, COL.displayName),
    userPrincipalName: upn,
    mail: cellOrNull(row, COL.mail),
    employeeId: cellOrNull(row, COL.bmsId),
    onPremisesExtensionAttributes: {
      extensionAttribute14: cellOrNull(row, COL.localHrId),
      extensionAttribute15: cellOrNull(row, COL.snTicketId),
    },
    givenName: cellOrNull(row, COL.firstName),
    surname: cellOrNull(row, COL.lastName),
    country: cellOrNull(row, COL.country),
    city: cellOrNull(row, COL.city),
    jobTitle: cellOrNull(row, COL.jobTitle),
    officeLocation: cellOrNull(row, COL.officeLocation),
    streetAddress: cellOrNull(row, COL.streetAddress),
    state: cellOrNull(row, COL.state),
    postalCode: cellOrNull(row, COL.postalCode),
    businessPhones: businessPhone ? [businessPhone] : [],
    mobilePhone: cellOrNull(row, COL.mobilePhone),
    companyName: cellOrNull(row, COL.companyName),
    department: cellOrNull(row, COL.department),
  };

  if (mailNickname) {
    body.mailNickname = mailNickname;
  }

  return body;
}

export async function updateUser(
  row: unknown[],
  accessToken: string
): Promise<UpdateUserResult> {
  const objectId = cell(row, COL.objectId);
  if (!objectId) {
    return { status: 'error', error: 'Object ID is required' };
  }

  const body = rowToGraphUpdateBody(row);

  try {
    const response = await fetch(
      `${GRAPH_USERS_URL}/${encodeURIComponent(objectId)}`,
      {
        method: 'PATCH',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(body),
      }
    );

    if (response.status === 204 || response.ok) {
      return { status: 'success' };
    }

    let errorMessage: string;
    const text = await response.text();
    try {
      const errJson = JSON.parse(text) as {
        error?: { message?: string; code?: string };
      };
      errorMessage =
        (errJson.error?.message ?? text) || `HTTP ${response.status}`;
    } catch {
      errorMessage = text || `HTTP ${response.status}`;
    }

    return { status: 'error', error: errorMessage };
  } catch (err) {
    const errorMessage = err instanceof Error ? err.message : String(err);
    return { status: 'error', error: errorMessage };
  }
}
