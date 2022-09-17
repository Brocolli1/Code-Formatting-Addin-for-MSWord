import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import {ButtonPrimaryExample} from './Button';

/* global Word, require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Design",
          primaryText: "Make your code look fresh with colors and formatting",
        },
      ],
    });
  }

  click = async () => {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

      await context.sync();
    });
  };

  render() {
    return (
      <div className="ms-welcome">
      <Header logo="assets/new_logo.png" title={this.props.title} message="Code in word" />
      <HeroList message="Write great looking code in MS Word with this Addin" items={this.state.listItems} >
        <ButtonPrimaryExample />
      </HeroList>
      </div>
    );
  }
}
