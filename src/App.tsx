import {
  Providers,
  SimpleProvider,
  ProviderState,
} from "@microsoft/mgt-element";
import Persons from "./components/Persons";
import { Switch } from "antd";
import "./App.css";
import PeoplePickerCmp from "./components/PeoplePickerCmp";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { MgtPerson } from "@microsoft/mgt-components";
import {
  Agenda,
  FileList,
  Get,
  MgtTemplateProps,
  Person,
  PersonCard,
  PersonViewType,
  TeamsChannelPicker,
} from "@microsoft/mgt-react";
import { useRef, useState } from "react";

Providers.globalProvider = new SimpleProvider(
  (_scopes: string[]): Promise<string> => {
    return new Promise((resolve) => {
      resolve(
        "eyJ0eXAiOiJKV1QiLCJub25jZSI6ImtRdmg0UVI1V2t0SjNUNjd2WTZvbTlhMVRja2lUZnVzaGpCNW8xdk5IT1kiLCJhbGciOiJSUzI1NiIsIng1dCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9kMjBjNWM4Yi1mZDQ5LTQ4ZTYtODNkNi03MmI1Nzc0MmMyYTMvIiwiaWF0IjoxNjkxNzQ0MTAwLCJuYmYiOjE2OTE3NDQxMDAsImV4cCI6MTY5MTc0NTYwMCwiYWlvIjoiRTJGZ1lEajQ2WWZmNUl0Nm0vemZybVM3L0l5SENRQT0iLCJhcHBfZGlzcGxheW5hbWUiOiJiaWthc2gtdGVzdC1hcHAiLCJhcHBpZCI6IjNlNzZjYzg0LTQ4NjQtNDVlNC05YTBiLTk5NjkxNWI3Y2JkOSIsImFwcGlkYWNyIjoiMSIsImlkcCI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0L2QyMGM1YzhiLWZkNDktNDhlNi04M2Q2LTcyYjU3NzQyYzJhMy8iLCJpZHR5cCI6ImFwcCIsIm9pZCI6ImU0MGYyMGFkLWNiMTEtNDQ1Zi05MTY1LWFjNDFhZWVmZjM4YiIsInJoIjoiMC5BWFlBaTF3TTBrbjk1a2lEMW5LMWQwTENvd01BQUFBQUFBQUF3QUFBQUFBQUFBQzBBQUEuIiwicm9sZXMiOlsiUHJlc2VuY2UuUmVhZFdyaXRlLkFsbCIsIlBlb3BsZS5SZWFkLkFsbCIsIlVzZXIuUmVhZC5BbGwiLCJHcm91cE1lbWJlci5SZWFkLkFsbCJdLCJzdWIiOiJlNDBmMjBhZC1jYjExLTQ0NWYtOTE2NS1hYzQxYWVlZmYzOGIiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiTkEiLCJ0aWQiOiJkMjBjNWM4Yi1mZDQ5LTQ4ZTYtODNkNi03MmI1Nzc0MmMyYTMiLCJ1dGkiOiJvVTNWX3JkaTlrcWhoSkJyakJDaUFBIiwidmVyIjoiMS4wIiwid2lkcyI6WyIwOTk3YTFkMC0wZDFkLTRhY2ItYjQwOC1kNWNhNzMxMjFlOTAiXSwieG1zX3RjZHQiOjE0NDY3OTY3NjV9.hHxyHGpD35kjqEV3sqysiQsYPiBlyj0-KoPTWY9mrrk07FzLVISZwAiskhZCyIkzSf9ulPDYS_1FkMzM0LSD_PWdP7JWlxjUdH6GKoYhI4RvGo1uqRPt3blX7os-kZdZM6psIh91jrRRIZ93hVqTThFBTu3AJcpa9rh9-k-d-Ov3U59hHBJjSCKvTNVuSkX7OPtSuGpAbTXOocYuOTL__SST30bEuRo-IAsMotv5j21J4GsXKyqvD2jpwc-B_jskW0A4T68N611R1MORtZOXLtpchptlS1UTmNn0KBQd_x8Audf-74J4Yom5V4kqL6fXD-l4gHStmlMHZMTrQk-59Q"
      );
    });
  }
);

Providers.globalProvider.setState(ProviderState.SignedIn);

function App() {
  const [theme, setTheme] = useState(false);

  const handleTemplateRendered = (e: Event) => {
    console.log("Event Rendered: ", e);
  };

  const MyTemplate = (props: MgtTemplateProps) => {
    const me = props.dataContext as MicrosoftGraph.User;
    return <div>hello {me.displayName}</div>;
  };

  const MyEvent = (props: MgtTemplateProps) => {
    const { event } = props.dataContext as { event: MicrosoftGraph.Event };
    return <div>{event.subject}</div>;
  };

  const MyMessage = (props: MgtTemplateProps) => {
    const message = props.dataContext as MicrosoftGraph.Message;

    const personRef = useRef<MgtPerson>();

    const handlePersonClick = () => {
      console.log(personRef.current);
    };

    return (
      <div>
        <b>Subject:</b>
        {message.subject}
        <div>
          <b>From:</b>
          <Person
            ref={personRef}
            onClick={handlePersonClick}
            personQuery={message.from?.emailAddress?.address || ""}
            fallbackDetails={{
              mail: message.from?.emailAddress?.address,
              displayName: message.from?.emailAddress?.name,
            }}
            view={PersonViewType.oneline}
          ></Person>
        </div>
      </div>
    );
  };

  return (
    <div
      className="rootcls"
      style={{ background: theme ? "#45657d" : "#f7f7f7" }}
    >
      <div style={{ float: "right", margin: "10px" }}>
        <span style={{ marginRight: "10px" }}>{theme ? "Dark" : "Light"}</span>
        <Switch
          defaultChecked
          onChange={() => {
            setTheme(!theme);
          }}
        />
      </div>
      <div className="App">
        <div className="person-cls">
          <Persons />
        </div>

        <div className="people-picker-cls">
          <PeoplePickerCmp />
        </div>

        <div className="people-picker-cls">
          <PersonCard
            showPresence
            isExpanded
            personQuery="me"
            userId="bikash.bagaria0@cfexperian.com"
          />
        </div>

        <Agenda groupByDay templateRendered={handleTemplateRendered}>
          <MyEvent template="event" />
        </Agenda>

        <Get resource="/me">
          <MyTemplate />
        </Get>

        <Get resource="/me/messages" scopes={["mail.read"]} maxPages={2}>
          <MyMessage template="value" />
        </Get>

        <TeamsChannelPicker />
      </div>
    </div>
  );
}

export default App;
