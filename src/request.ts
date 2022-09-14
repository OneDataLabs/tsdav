import { fetch } from 'cross-fetch';
import getLogger from 'debug';
import convert, { ElementCompact } from 'xml-js';

import { DAVNamespace, DAVNamespaceShort } from './consts';
import { DAVDepth, DAVRequest, DAVResponse } from './types/DAVTypes';
import { DAVAccount } from './types/models';
import { camelCase } from './util/camelCase';
import { nativeType } from './util/nativeType';
import { cleanupFalsy, getDAVAttribute } from './util/requestHelpers';

const DigestFetch = require('digest-fetch');

const debug = getLogger('tsdav:request');

type RawProp = { prop: { [key: string]: any }; status: string; responsedescription?: string };
type RawResponse = {
  href: string;
  status: string;
  ok: boolean;
  error: { [key: string]: any };
  responsedescription: string;
  propstat: RawProp | RawProp[];
};

export const davRequest = async (params: {
  url: string;
  init: DAVRequest;
  account?: DAVAccount;
  convertIncoming?: boolean;
  parseOutgoing?: boolean;
}): Promise<DAVResponse[]> => {
  const { url, init, account, convertIncoming = true, parseOutgoing = true } = params;
  const { headers, body, namespace, method, attributes } = init;
  const xmlBody = convertIncoming
    ? convert.js2xml(
        {
          _declaration: { _attributes: { version: '1.0', encoding: 'utf-8' } },
          ...body,
          _attributes: attributes,
        },
        {
          compact: true,
          spaces: 2,
          elementNameFn: (name) => {
            // add namespace to all keys without namespace
            if (namespace && !/^.+:.+/.test(name)) {
              return `${namespace}:${name}`;
            }
            return name;
          },
        }
      )
    : body;

  // For Digest auth
  const client = new DigestFetch(account!.credentials!.username, account!.credentials!.password, { algorithm: 'MD5' });
  const davResponse = await client.fetch(url, {
    headers: {
      'Content-Type': 'application/xml; charset=utf-8'
    },
    body: xmlBody,
    method
  })

  const resText = await davResponse.text();

  if (
    !davResponse.ok ||
    !davResponse.headers.get('content-type')?.includes('xml') ||
    !parseOutgoing
  ) {
    return [
      {
        href: davResponse.url,
        ok: davResponse.ok,
        status: davResponse.status,
        statusText: davResponse.statusText,
        raw: resText,
      },
    ];
  }

  const result: any = convert.xml2js(resText, {
    compact: true,
    trim: true,
    textFn: (value: any, parentElement: any) => {
      try {
        // This is needed for xml-js design reasons
        // eslint-disable-next-line no-underscore-dangle
        const parentOfParent = parentElement._parent;
        const pOpKeys = Object.keys(parentOfParent);
        const keyNo = pOpKeys.length;
        const keyName = pOpKeys[keyNo - 1];
        const arrOfKey = parentOfParent[keyName];
        const arrOfKeyLen = arrOfKey.length;
        if (arrOfKeyLen > 0) {
          const arr = arrOfKey;
          const arrIndex = arrOfKey.length - 1;
          arr[arrIndex] = nativeType(value);
        } else {
          parentOfParent[keyName] = nativeType(value);
        }
      } catch (e) {
        debug((e as Error).stack);
      }
    },
    // remove namespace & camelCase
    elementNameFn: (attributeName) => camelCase(attributeName.replace(/^.+:/, '')),
    attributesFn: (value: any) => {
      const newVal = { ...value };
      delete newVal.xmlns;
      return newVal;
    },
    ignoreDeclaration: true,
  });

  const responseBodies: RawResponse[] = Array.isArray(result.multistatus.response)
    ? result.multistatus.response
    : [result.multistatus.response];

  return responseBodies.map((responseBody) => {
    const statusRegex = /^\S+\s(?<status>\d+)\s(?<statusText>.+)$/;
    if (!responseBody) {
      return {
        status: davResponse.status,
        statusText: davResponse.statusText,
        ok: davResponse.ok,
      };
    }

    const matchArr = statusRegex.exec(responseBody.status);

    return {
      raw: result,
      href: responseBody.href,
      status: matchArr?.groups ? Number.parseInt(matchArr?.groups.status, 10) : davResponse.status,
      statusText: matchArr?.groups?.statusText ?? davResponse.statusText,
      ok: !responseBody.error,
      error: responseBody.error,
      responsedescription: responseBody.responsedescription,
      props: (Array.isArray(responseBody.propstat)
        ? responseBody.propstat
        : [responseBody.propstat]
      ).reduce((prev, curr) => {
        return {
          ...prev,
          ...curr?.prop,
        };
      }, {}),
    };
  });
};

export const propfind = async (params: {
  url: string;
  props: ElementCompact;
  depth?: DAVDepth;
  headers?: Record<string, string>;
  account?: DAVAccount;
}): Promise<DAVResponse[]> => {
  const { url, props, depth, headers, account } = params;
  return davRequest({
    url,
    init: {
      method: 'PROPFIND',
      headers: cleanupFalsy({ depth, ...headers }),
      namespace: DAVNamespaceShort.DAV,
      body: {
        propfind: {
          _attributes: getDAVAttribute([
            DAVNamespace.CALDAV,
            DAVNamespace.CALDAV_APPLE,
            DAVNamespace.CALENDAR_SERVER,
            DAVNamespace.CARDDAV,
            DAVNamespace.DAV,
          ]),
          prop: props,
        },
      },
    },
    account
  });
};

export const createObject = async (params: {
  url: string;
  data: BodyInit;
  headers?: Record<string, string>;
  account?: DAVAccount;
}): Promise<Response> => {
  const { url, data, headers, account } = params;
  const client = new DigestFetch(account!.credentials!.username, account!.credentials!.password, { algorithm: 'MD5' });
  return client.fetch(url, { method: 'PUT', body: data, headers });
};

export const updateObject = async (params: {
  url: string;
  data: BodyInit;
  etag?: string;
  headers?: Record<string, string>;
  account?: DAVAccount;
}): Promise<Response> => {
  const { url, data, etag, headers, account } = params;
  const client = new DigestFetch(account!.credentials!.username, account!.credentials!.password, { algorithm: 'MD5' });
  return client.fetch(url, {
    method: 'PUT',
    body: data,
    headers: cleanupFalsy({ 'If-Match': etag, ...headers }),
  });
};

export const deleteObject = async (params: {
  url: string;
  etag?: string;
  headers?: Record<string, string>;
  account?: DAVAccount;
}): Promise<Response> => {
  const { url, headers, etag, account } = params;
  const client = new DigestFetch(account!.credentials!.username, account!.credentials!.password, { algorithm: 'MD5' });
  return client.fetch(url, {
    method: 'DELETE',
    headers: cleanupFalsy({ 'If-Match': etag, ...headers }),
  });
};