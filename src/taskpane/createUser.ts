/**
 * Create User via Microsoft Graph API (POST /users).
 * Maps CreateUsers table columns to Graph user properties.
 */

import {
  GRAPH_USERS_URL,
  cell,
  parseGraphErrorResponse,
  toErrorMessage,
} from './graphHelpers';

export interface CreateUserResult {
  objectId: string | null;
  generatedPassword: string | null;
  status: 'success' | 'error' | 'skipped';
  error?: string;
}

/** Column indices matching USERS_TABLE_HEADERS in App.tsx */
const COL = {
  userPrincipalName: 0,
  mail: 1,
  bmsId: 2,
  localHrId: 3,
  snTicketId: 4,
  firstName: 5,
  lastName: 6,
  displayName: 7,
  country: 8,
  city: 9,
  jobTitle: 10,
  officeLocation: 11,
  streetAddress: 12,
  state: 13,
  postalCode: 14,
  businessPhone: 15,
  mobilePhone: 16,
  companyName: 17,
  department: 18,
} as const;

function generatePassword(): string {
  const length = 16;
  const lower = 'abcdefghjkmnpqrstuvwxyz';
  const upper = 'ABCDEFGHJKMNPQRSTUVWXYZ';
  const digits = '23456789';
  const special = '!@#$%&*';
  const all = lower + upper + digits + special;
  const getRandom = (chars: string) =>
    chars[Math.floor(Math.random() * chars.length)];
  let pw =
    getRandom(lower) +
    getRandom(upper) +
    getRandom(digits) +
    getRandom(special);
  for (let i = pw.length; i < length; i++) {
    pw += all[Math.floor(Math.random() * all.length)];
  }
  return pw
    .split('')
    .sort(() => Math.random() - 0.5)
    .join('');
}

interface GraphUserBody {
  accountEnabled: boolean;
  displayName: string;
  mailNickname: string;
  userPrincipalName: string;
  passwordProfile: {
    password: string;
    forceChangePasswordNextSignIn: boolean;
  };
  mail?: string;
  employeeId?: string;
  onPremisesExtensionAttributes?: {
    extensionAttribute14?: string;
    extensionAttribute15?: string;
  };
  givenName?: string;
  surname?: string;
  country?: string;
  city?: string;
  jobTitle?: string;
  officeLocation?: string;
  streetAddress?: string;
  state?: string;
  postalCode?: string;
  businessPhones?: string[];
  mobilePhone?: string;
  companyName?: string;
  department?: string;
}

function rowToGraphUser(row: unknown[], password: string): GraphUserBody {
  const upn = cell(row, COL.userPrincipalName);
  const mailNickname = upn.includes('@') ? upn.split('@')[0]! : upn;

  const body: GraphUserBody = {
    accountEnabled: true,
    displayName: cell(row, COL.displayName),
    mailNickname,
    userPrincipalName: upn,
    passwordProfile: {
      password,
      forceChangePasswordNextSignIn: true,
    },
  };

  const mail = cell(row, COL.mail);
  if (mail) body.mail = mail;

  const employeeId = cell(row, COL.bmsId);
  if (employeeId) body.employeeId = employeeId;

  const localHrId = cell(row, COL.localHrId);
  const snTicketId = cell(row, COL.snTicketId);
  if (localHrId || snTicketId) {
    body.onPremisesExtensionAttributes = {};
    if (localHrId) body.onPremisesExtensionAttributes.extensionAttribute14 = localHrId;
    if (snTicketId) body.onPremisesExtensionAttributes.extensionAttribute15 = snTicketId;
  }

  const givenName = cell(row, COL.firstName);
  if (givenName) body.givenName = givenName;

  const surname = cell(row, COL.lastName);
  if (surname) body.surname = surname;

  const country = cell(row, COL.country);
  if (country) body.country = country;

  const city = cell(row, COL.city);
  if (city) body.city = city;

  const jobTitle = cell(row, COL.jobTitle);
  if (jobTitle) body.jobTitle = jobTitle;

  const officeLocation = cell(row, COL.officeLocation);
  if (officeLocation) body.officeLocation = officeLocation;

  const streetAddress = cell(row, COL.streetAddress);
  if (streetAddress) body.streetAddress = streetAddress;

  const state = cell(row, COL.state);
  if (state) body.state = state;

  const postalCode = cell(row, COL.postalCode);
  if (postalCode) body.postalCode = postalCode;

  const businessPhone = cell(row, COL.businessPhone);
  if (businessPhone) body.businessPhones = [businessPhone];

  const mobilePhone = cell(row, COL.mobilePhone);
  if (mobilePhone) body.mobilePhone = mobilePhone;

  const companyName = cell(row, COL.companyName);
  if (companyName) body.companyName = companyName;

  const department = cell(row, COL.department);
  if (department) body.department = department;

  return body;
}

/** Check if a user with the given userPrincipalName already exists (Graph API). */
async function userExistsByUpn(
  userPrincipalName: string,
  accessToken: string
): Promise<boolean> {
  const encoded = encodeURIComponent(`userPrincipalName eq '${userPrincipalName.replace(/'/g, "''")}'`);
  const url = `${GRAPH_USERS_URL}?$filter=${encoded}&$select=id&$top=1`;
  const response = await fetch(url, {
    method: 'GET',
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!response.ok) return false;
  const data = (await response.json()) as { value?: unknown[] };
  return Array.isArray(data.value) && data.value.length > 0;
}

export async function createUser(
  row: unknown[],
  accessToken: string
): Promise<CreateUserResult> {
  const upn = cell(row, COL.userPrincipalName);
  if (!upn) {
    return {
      objectId: null,
      generatedPassword: null,
      status: 'error',
      error: 'userPrincipalName is required',
    };
  }

  try {
    const exists = await userExistsByUpn(upn, accessToken);
    if (exists) {
      return {
        objectId: null,
        generatedPassword: null,
        status: 'skipped',
        error: 'UPN already exists',
      };
    }
  } catch (err) {
    console.warn('userExistsByUpn check failed, continuing with create:', toErrorMessage(err));
  }

  const password = generatePassword();
  const body = rowToGraphUser(row, password);

  try {
    const response = await fetch(GRAPH_USERS_URL, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(body),
    });

    if (response.status === 201) {
      const data = (await response.json()) as { id?: string };
      return {
        objectId: data.id ?? null,
        generatedPassword: password,
        status: 'success',
      };
    }

    const errorMessage = await parseGraphErrorResponse(response);
    return {
      objectId: null,
      generatedPassword: null,
      status: 'error',
      error: errorMessage,
    };
  } catch (err) {
    return {
      objectId: null,
      generatedPassword: null,
      status: 'error',
      error: toErrorMessage(err),
    };
  }
}
