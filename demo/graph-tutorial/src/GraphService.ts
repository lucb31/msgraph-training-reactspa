// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <graphServiceSnippet1>
import moment, {Moment} from 'moment';
import {Event, Group, User} from 'microsoft-graph';
import {GraphRequestOptions, PageCollection, PageIterator} from '@microsoft/microsoft-graph-client';

var graph = require('@microsoft/microsoft-graph-client');

function getAuthenticatedClient(accessToken: string) {
  // Initialize Graph client
  const client = graph.Client.init({
    // Use the provided access token to authenticate
    // requests
    authProvider: (done: any) => {
      done(null, accessToken);
    }
  });

  return client;
}

export async function getUserDetails(accessToken: string) {
  const client = getAuthenticatedClient(accessToken);

  const user = await client
    .api('/me')
    .select('displayName,mail,mailboxSettings,userPrincipalName')
    .get();

  if (user.value)
    return user.value;
  else
    return user;
}
// </graphServiceSnippet1>

// <getUserWeekCalendarSnippet>
export async function getUserWeekCalendar(accessToken: string, timeZone: string, startDate: Moment): Promise<Event[]> {
  const client = getAuthenticatedClient(accessToken);

  // Generate startDateTime and endDateTime query params
  // to display a 7-day window
  var startDateTime = startDate.format();
  var endDateTime = moment(startDate).add(7, 'day').format();

  // GET /me/calendarview?startDateTime=''&endDateTime=''
  // &$select=subject,organizer,start,end
  // &$orderby=start/dateTime
  // &$top=50
  var response: PageCollection = await client
    .api('/me/calendarview')
    .header('Prefer', `outlook.timezone="${timeZone}"`)
    .query({ startDateTime: startDateTime, endDateTime: endDateTime })
    .select('subject,organizer,start,end')
    .orderby('start/dateTime')
    .top(25)
    .get();

  if (response["@odata.nextLink"]) {
    // Presence of the nextLink property indicates more results are available
    // Use a page iterator to get all results
    var events: Event[] = [];

    // Must include the time zone header in page
    // requests too
    var options: GraphRequestOptions = {
      headers: { 'Prefer': `outlook.timezone="${timeZone}"` }
    };

    var pageIterator = new PageIterator(client, response, (event) => {
      events.push(event);
      return true;
    }, options);

    await pageIterator.iterate();

    return events;
  } else {

    return response.value;
  }
}
// </getUserWeekCalendarSnippet>

// <createEventSnippet>
export async function createEvent(accessToken: string, newEvent: Event): Promise<Event> {
  const client = getAuthenticatedClient(accessToken);

  // POST /me/events
  // JSON representation of the new event is sent in the
  // request body
  return await client
    .api('/me/events')
    .post(newEvent);
}
// </createEventSnippet>

export async function createGroup(accessToken: string, newGroup: Group): Promise<Group> {
  const client = getAuthenticatedClient(accessToken);
  const response = await client
      .api('/groups')
      .post(newGroup);
  if (response["value"])
    return response["value"];
  else
    return response;
}

export async function addMembersToGroup(accessToken: string, group: Group, members: User[]): Promise<Group> {
  const client = getAuthenticatedClient(accessToken);

  const bindArray:string[] = [];
  members.forEach(member => bindArray.push('https://graph.microsoft.com/v1.0/directoryObjects/' + member.id));
  const updateGroup = {
    "members@odata.bind": bindArray
  };
  return await client
      .api("/groups/" + group.id)
      .update(updateGroup);
/*
  return await client
      .api("/groups/" + group.id + "/members/$ref")
      .post(member);
 */
}

export async function getGroups(accessToken: string): Promise<Group []> {
  const client = getAuthenticatedClient(accessToken);

  return await client
      .api('/groups')
      .get();
}

export async function getUsers(accessToken: string): Promise<User []> {
  const client = getAuthenticatedClient(accessToken);
  const response = await client
      .api('/users')
      .get();

  if (response["value"])
    return response["value"];
  else
    return response;
}
