// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
import React from 'react';
import { NavLink as RouterNavLink } from 'react-router-dom';
import { Table } from 'reactstrap';
import moment, { Moment } from 'moment-timezone';
import { findOneIana } from "windows-iana";
import { Event } from 'microsoft-graph';
import { config } from './Config';
import { getUserWeekCalendar } from './GraphService';
import withAuthProvider, { AuthComponentProps } from './AuthProvider';
import CalendarDayRow from './CalendarDayRow';
import './Calendar.css';

// MGT
import { Providers, SimpleProvider } from '@microsoft/mgt';
import { Agenda, MgtTemplateProps } from '@microsoft/mgt-react';

interface CalendarState {
  eventsLoaded: boolean;
  events: Event[];
  startOfWeek: Moment | undefined;
}

const MyEvent = (props: MgtTemplateProps) => {
  const { event } = props.dataContext;
  return <div>{event.subject}</div>;
};

class Calendar extends React.Component<AuthComponentProps, CalendarState> {
  mgtProvider: SimpleProvider = new SimpleProvider(async (scopes: string[]) => {
    return await this.props.getAccessToken(scopes);
  });

  constructor(props: any) {
    super(props);

    Providers.globalProvider = this.mgtProvider;
    this.state = {
      eventsLoaded: false,
      events: [],
      startOfWeek: undefined
    };
  }

  async componentDidUpdate()
  {
    if (this.props.user && !this.state.eventsLoaded)
    {
      try {
        // Get the user's access token
        var accessToken = await this.props.getAccessToken(config.scopes);

        // Convert user's Windows time zone ("Pacific Standard Time")
        // to IANA format ("America/Los_Angeles")
        // Moment needs IANA format
        var ianaTimeZone = findOneIana(this.props.user.timeZone);

        // Get midnight on the start of the current week in the user's timezone,
        // but in UTC. For example, for Pacific Standard Time, the time value would be
        // 07:00:00Z
        var startOfWeek = moment.tz(ianaTimeZone!.valueOf()).startOf('week').utc();

        // Get the user's events
        var events = await getUserWeekCalendar(accessToken, this.props.user.timeZone, startOfWeek);

        // Update the array of events in state
        this.setState({
          eventsLoaded: true,
          events: events,
          startOfWeek: startOfWeek
        });
      }
      catch (err) {
        this.props.setError('ERROR', JSON.stringify(err));
      }
    }
  }

  // <renderSnippet>
  render() {
    var sunday = moment(this.state.startOfWeek);
    var monday = moment(sunday).add(1, 'day');
    var tuesday = moment(monday).add(1, 'day');
    var wednesday = moment(tuesday).add(1, 'day');
    var thursday = moment(wednesday).add(1, 'day');
    var friday = moment(thursday).add(1, 'day');
    var saturday = moment(friday).add(1, 'day');

    // Agenda is in mgt-react package, which isn't mentioned in docs
    return (
      <div>
        <div className="mb-3">
          <h1 className="mb-3">{sunday.format('MMMM D, YYYY')} - {saturday.format('MMMM D, YYYY')}</h1>
          <RouterNavLink to="/newevent" className="btn btn-light btn-sm" exact>New event</RouterNavLink>
        </div>
        <Agenda
          events={this.state.events}
          groupByDay={true} />
      </div>
    );
  }
  // </renderSnippet>
}

export default withAuthProvider(Calendar);
