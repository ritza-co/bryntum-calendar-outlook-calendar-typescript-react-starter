import { useRef, useEffect, useState, useCallback } from 'react';
import { AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react';
import { findIana } from 'windows-iana';
import bryntumLogo from '../assets/bryntum-symbol-white.svg';
import { BryntumButton, BryntumCalendar } from '@bryntum/calendar-react';
import { createCalendarConfig } from '../calendarConfig';
import { useAppContext } from '../AppContext';
import SignInModal from './SignInModal';
import { EventModel, Toast } from '@bryntum/calendar';
import { getUserFutureCalendar, getUserPastCalendar, getUserWeekCalendar } from '../graphService';
import { BryntumSync } from '../crudHelpers';
import { SyncDataParams } from '../types';

export default function Calendar() {
    const calendarRef = useRef(null);
    const app = useAppContext();
    const [events, setEvents] = useState<Partial<EventModel>[]>();
    const hasFetchedInitialEvents = useRef(false);
    const hasRunFirstEffect = useRef(false);
    const [timeZone, setTimeZone] = useState('UTC');

    useEffect(() => {
        if (app.user?.timeZone) {
            const ianaTimeZones = findIana(app.user.timeZone);
            setTimeZone(ianaTimeZones[0].valueOf());
        }
    }, [app.user?.timeZone]);

    const syncWithOutlook = useCallback((
        id: string,
        name: string,
        startDate: string,
        endDate: string,
        allDay: boolean,
        action: 'add' | 'update' | 'remove' | 'removeAll' | 'clearchanges' | 'filter' | 'update' | 'dataset' | 'replace'
    ) => {
        if (!app.authProvider) return;
        return BryntumSync(
            id,
            name,
            startDate,
            endDate,
            allDay,
            action,
            setEvents,
            app.authProvider,
            timeZone
        );
    }, [app.authProvider, timeZone]);

    const syncData = useCallback(({ action, records }: SyncDataParams) => {
        if ((action === 'add' && !records[0].copyOf) || action === 'dataset') {
            return;
        }
        records.forEach((record) => {
            syncWithOutlook(
                record.get('id'),
                record.get('name'),
                record.get('startDate'),
                record.get('endDate'),
                record.get('allDay'),
                action
            );
        });
    }, [syncWithOutlook]);

    const addRecord = useCallback((event: { eventRecord: EventModel }) => {
        const { eventRecord } = event;
        const isNew = eventRecord.id.toString().startsWith('_generated');

        // Get the date strings in the calendar's timezone
        const startDate = new Date(eventRecord.startDate);
        const endDate = new Date(eventRecord.endDate);

        syncWithOutlook(
            eventRecord.id.toString(),
            eventRecord.name || '',
            startDate.toISOString(),
            endDate.toISOString(),
            eventRecord.allDay || false,
            isNew ? 'add' : 'update'
        );
    }, [syncWithOutlook]);

    useEffect(() => {
        const loadEvents = async() => {
            if (app.user && !events) {
                if (hasRunFirstEffect.current) {
                    return;
                }
                hasRunFirstEffect.current = true;
                try {
                    const ianaTimeZones = findIana(app.user?.timeZone || 'UTC');
                    const outlookEvents = await getUserWeekCalendar(app.authProvider!, ianaTimeZones[0].valueOf());
                    const calendarEvents: Partial<EventModel>[] = [];
                    outlookEvents.forEach((event) => {
                        // Convert the dates to the calendar's timezone
                        const startDate = event.start?.dateTime ? new Date(event.start.dateTime) : null;
                        const endDate = event.end?.dateTime ? new Date(event.end.dateTime) : null;

                        calendarEvents.push({
                            id        : `${event.id}`,
                            name      : `${event.subject}`,
                            startDate : startDate?.toISOString(),
                            endDate   : endDate?.toISOString(),
                            allDay    : event.isAllDay || false
                        });
                    });
                    setEvents(calendarEvents);
                }
                catch (err) {
                    const error = err as Error;
                  app.displayError!(error.message);
                }
            }
        };

        loadEvents();
    }, [app.user, app.authProvider, events, app.displayError]);

    // load more events from outlook
    useEffect(() => {
        // Skip if still loading or if we've already fetched events
        if (!app.user) return;
        if (app.isLoading) return;
        if (hasFetchedInitialEvents.current) return;

        async function fetchAllData() {
            try {
                hasFetchedInitialEvents.current = true;
                const ianaTimeZones = findIana(app.user?.timeZone || 'UTC');
                const [calendarPastEvents, calendarFutureEvents] = await Promise.all([getUserPastCalendar(app.authProvider!, ianaTimeZones[0].valueOf()), getUserFutureCalendar(app.authProvider!, ianaTimeZones[0].valueOf())]);

                const pastEvents: Partial<EventModel>[] = [];
                const futureEvents: Partial<EventModel>[] = [];

                calendarPastEvents.forEach((event) => {
                    // Convert the dates to the calendar's timezone
                    const startDate = event.start?.dateTime ? new Date(event.start.dateTime) : null;
                    const endDate = event.end?.dateTime ? new Date(event.end.dateTime) : null;

                    pastEvents.push({
                        id        : `${event.id}`,
                        name      : `${event.subject}`,
                        startDate : startDate?.toISOString(),
                        endDate   : endDate?.toISOString(),
                        allDay    : event.isAllDay || false
                    });
                });

                calendarFutureEvents.forEach((event) => {
                    // Convert the dates to the calendar's timezone
                    const startDate = event.start?.dateTime ? new Date(event.start.dateTime) : null;
                    const endDate = event.end?.dateTime ? new Date(event.end.dateTime) : null;

                    futureEvents.push({
                        id        : `${event.id}`,
                        name      : `${event.subject}`,
                        startDate : startDate?.toISOString(),
                        endDate   : endDate?.toISOString(),
                        allDay    : event.isAllDay || false
                    });
                });

                setEvents(currentEvents => {
                    return [...pastEvents, ...(currentEvents || []), ...futureEvents];
                });
            }
            catch (err) {
                const error = err as Error;
              app.displayError!(error.message);
            }
        }
        fetchAllData();
    }, [app.user, app.authProvider, app.displayError, app.isLoading, app.user?.timeZone]);


    useEffect(() => {
        if (app.error) {
            Toast.show({
                html    : app.error.message,
                timeout : 0
            });
        }
    }, [app.error]);


    const calendarProps = createCalendarConfig({ syncData, addRecord });

    return (
        <>
            <UnauthenticatedTemplate>
                <SignInModal />
            </UnauthenticatedTemplate>
            <header>
                <div className="title-container">
                    <img src={bryntumLogo} role="presentation" />
                    <h1>
                Bryntum Calendar synced with Outlook Calendar demo
                    </h1>
                </div>
                <AuthenticatedTemplate>
                    <BryntumButton
                        cls="b-raised"
                        text={app.user && app.isLoading ? 'Signing out...' : 'Sign out'}
                        color='b-blue'
                        onClick={() => app.signOut?.()}
                        disabled={app.isLoading}
                    />
                </AuthenticatedTemplate>
                <UnauthenticatedTemplate>
                    <BryntumButton
                        cls="b-raised"
                        text={app.isLoading ? 'Signing in...' : 'Sign in with Microsoft'}
                        color='b-blue'
                        onClick={() => app.signIn?.()}
                        disabled={app.isLoading}
                    />
                </UnauthenticatedTemplate>
            </header>
            <BryntumCalendar
                ref={calendarRef}
                eventStore={{
                    data : events
                }}
                {...calendarProps}
            />
        </>
    );
};