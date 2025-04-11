import { Client, GraphRequestOptions, PageCollection, PageIterator
} from '@microsoft/microsoft-graph-client';
import { AuthCodeMSALBrowserAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser';
import { endOfWeek, startOfWeek } from 'date-fns';
import { fromZonedTime } from 'date-fns-tz';
import { User, Event } from '@microsoft/microsoft-graph-types';

let graphClient: Client | undefined = undefined;

function ensureClient(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
    if (!graphClient) {
        graphClient = Client.initWithMiddleware({
            authProvider : authProvider
        });
    }
    return graphClient;
}

export async function getUser(authProvider: AuthCodeMSALBrowserAuthenticationProvider): Promise<User> {
    ensureClient(authProvider);

    // Return the /me API endpoint result as a User object
    const user: User = await graphClient!.api('/me')
    // Only retrieve the specific fields needed
        .select('displayName,mail,mailboxSettings,userPrincipalName')
        .get();
    return user;
}

export async function getUserWeekCalendar(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    timeZone: string): Promise<Event[]> {
    ensureClient(authProvider);

    // Generate startDateTime and endDateTime query params
    // to display a 7-day window
    const now = new Date();
    const startDateTime = fromZonedTime(startOfWeek(now), timeZone).toISOString();
    const endDateTime = fromZonedTime(endOfWeek(now), timeZone).toISOString();

    const response: PageCollection = await graphClient!
        .api('/me/calendarview')
        .header('Prefer', `outlook.timezone="${timeZone}"`)
        .query({ startDateTime : startDateTime, endDateTime : endDateTime })
        .select('id, subject, start, end, isAllDay')
        .orderby('start/dateTime')
        .top(1000)
        .get();

    if (response['@odata.nextLink']) {
    // Presence of the nextLink property indicates more results are available
    // Use a page iterator to get all results
        const events: Event[] = [];

        // Must include the time zone header in page
        // requests too
        const options: GraphRequestOptions = {
            headers : { 'Prefer' : `outlook.timezone="${timeZone}"` }
        };

        const pageIterator = new PageIterator(graphClient!, response, (event) => {
            events.push(event);
            return true;
        }, options);

        await pageIterator.iterate();

        return events;
    }
    else {

        return response.value;
    }
}

export async function getUserPastCalendar(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    timeZone: string,
    daysInPast: number = 365): Promise<Event[]> {
    ensureClient(authProvider);

    // Calculate the date range
    const now = new Date();
    const endDateTime = fromZonedTime(startOfWeek(now), timeZone).toISOString();
    const startDateTime = new Date(now.getTime() - (daysInPast * 24 * 60 * 60 * 1000)).toISOString();

    // GET /me/calendarview with pagination support
    const response: PageCollection = await graphClient!
        .api('/me/calendarview')
        .header('Prefer', `outlook.timezone="${timeZone}"`)
        .query({ startDateTime, endDateTime })
        .select('id, subject, start, end, isAllDay')
        .orderby('start/dateTime')
        .top(1000)  // Maximum events per page
        .get();

    // If there are more than 1000 events, handle pagination
    if (response['@odata.nextLink']) {
        const events: Event[] = [];

        const options: GraphRequestOptions = {
            headers : { 'Prefer' : `outlook.timezone="${timeZone}"` }
        };

        // Use PageIterator to automatically handle fetching all pages
        const pageIterator = new PageIterator(graphClient!, response, (event) => {
            events.push(event);
            return true;
        }, options);

        await pageIterator.iterate();
        return events;
    }

    return response.value;
}

export async function getUserFutureCalendar(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    timeZone: string,
    daysInFuture: number = 365): Promise<Event[]> {
    ensureClient(authProvider);

    // Calculate the date range
    const now = new Date();
    const startDateTime = fromZonedTime(endOfWeek(now), timeZone).toISOString();
    const endDateTime = new Date(endOfWeek(now).getTime() + (daysInFuture * 24 * 60 * 60 * 1000)).toISOString();

    // GET /me/calendarview with pagination support
    const response: PageCollection = await graphClient!
        .api('/me/calendarview')
        .header('Prefer', `outlook.timezone="${timeZone}"`)
        .query({ startDateTime, endDateTime })
        .select('id, subject, start, end, isAllDay')
        .orderby('start/dateTime')
        .top(1000)  // Maximum events per page
        .get();

    // If there are more than 1000 events, handle pagination
    if (response['@odata.nextLink']) {
        const events: Event[] = [];

        const options: GraphRequestOptions = {
            headers : { 'Prefer' : `outlook.timezone="${timeZone}"` }
        };

        // Use PageIterator to automatically handle fetching all pages
        const pageIterator = new PageIterator(graphClient!, response, (event) => {
            events.push(event);
            return true;
        }, options);

        await pageIterator.iterate();
        return events;
    }

    return response.value;
}

export async function createEvent(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    newEvent: Event): Promise<Event> {
    ensureClient(authProvider);

    // POST /me/events
    // JSON representation of the new event is sent in the
    // request body
    return await graphClient!
        .api('/me/events')
        .post(newEvent);
}

export async function updateEvent(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    id: string,
    event: Event): Promise<Event> {
    ensureClient(authProvider);

    // POST /me/events
    // JSON representation of the new event is sent in the
    // request body
    return await graphClient!
        .api(`/me/events/${id}`)
        .patch(event);
}

export async function deleteEvent(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    id: string): Promise<Event> {
    ensureClient(authProvider);

    return await graphClient!
        .api(`/me/events/${id}`)
        .delete();
}