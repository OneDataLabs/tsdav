import URL from 'url';
import { DAVAccount, DAVAddressBook, DAVVCard } from 'models';
import getLogger from 'debug';
import { DAVProp, DAVDepth, DAVResponse } from 'DAVTypes';
import { DAVNamespace } from './consts';
import { collectionQuery, supportedReportSet } from './collection';
import { createObject, deleteObject, propfind, updateObject } from './request';
import { getDAVAttribute, formatProps, urlEquals } from './util/requestHelpers';

const debug = getLogger('tsdav:addressBook');

export const addressBookQuery = async (
  url: string,
  props: DAVProp[],
  options?: { depth?: DAVDepth; headers?: { [key: string]: any } }
): Promise<DAVResponse[]> => {
  return collectionQuery(
    url,
    {
      'addressbook-query': {
        _attributes: getDAVAttribute([DAVNamespace.CARDDAV, DAVNamespace.DAV]),
        prop: formatProps(props),
        filter: {
          'prop-filter': {
            _attributes: {
              name: 'FN',
            },
          },
        },
      },
    },
    { depth: options.depth, headers: options.headers }
  );
};

export const fetchAddressBooks = async (
  account: DAVAccount,
  options?: { headers?: { [key: string]: any } }
): Promise<DAVAddressBook[]> => {
  const res = await propfind(
    account.homeUrl,
    [
      { name: 'displayname', namespace: DAVNamespace.DAV },
      { name: 'getctag', namespace: DAVNamespace.CALENDAR_SERVER },
      { name: 'resourcetype', namespace: DAVNamespace.DAV },
      { name: 'sync-token', namespace: DAVNamespace.DAV },
    ],
    { depth: '1', headers: options.headers }
  );
  return Promise.all(
    res
      .filter((r) => r.props.displayname && r.props.displayname.length)
      .map((rs) => {
        debug(`Found address book named ${rs.props.displayname},
             props: ${JSON.stringify(rs.props)}`);
        return {
          data: rs,
          account,
          url: URL.resolve(account.rootUrl, rs.href),
          ctag: rs.props.getctag,
          displayName: rs.props.displayname,
          resourcetype: rs.props.resourcetype,
          syncToken: rs.props.syncToken,
        };
      })
      .map(async (addr) => ({ ...addr, reports: await supportedReportSet(addr, options) }))
  );
};

export const fetchVCards = async (
  addressBook: DAVAddressBook,
  options?: { headers?: { [key: string]: any } }
): Promise<DAVVCard[]> => {
  return (
    await addressBookQuery(
      addressBook.url,
      [
        { name: 'getetag', namespace: DAVNamespace.DAV },
        { name: 'address-data', namespace: DAVNamespace.CARDDAV },
      ],
      { depth: '1', headers: options.headers }
    )
  ).map((res) => {
    return {
      data: res,
      addressBook,
      url: URL.resolve(addressBook.account.rootUrl, res.href),
      etag: res.props.getetag,
      addressData: res.props.addressData,
    };
  });
};

export const createVCard = async (
  addressBook: DAVAddressBook,
  vCardString: string,
  filename: string,
  options?: { headers?: { [key: string]: any } }
): Promise<Response> => {
  return createObject(URL.resolve(addressBook.url, filename), vCardString, {
    headers: {
      'content-type': 'text/calendar; charset=utf-8',
      ...options.headers,
    },
  });
};

export const updateVCard = async (
  vCard: DAVVCard,
  options?: { headers?: { [key: string]: any } }
): Promise<Response> => {
  return updateObject(vCard.url, vCard.addressData, vCard.etag, {
    headers: {
      'content-type': 'text/calendar; charset=utf-8',
      ...options.headers,
    },
  });
};

export const deleteVCard = async (
  vCard: DAVVCard,
  options?: { headers?: { [key: string]: any } }
): Promise<Response> => {
  return deleteObject(vCard.url, vCard.etag, options);
};

// remote change -> local
export const syncCardDAVAccount = async (
  account: DAVAccount,
  options?: { headers?: { [key: string]: any } }
): Promise<DAVAccount> => {
  // find new Address books
  // TODO : implement syncCardDAVAccount
  const newAddressBooks = (await fetchAddressBooks(account, options)).filter((addr) =>
    account.addressBooks?.some((a) => urlEquals(a.url, addr.url))
  );
  return { accountType: 'carddav' };
};
