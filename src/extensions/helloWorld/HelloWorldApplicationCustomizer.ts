import * as React from 'react';
import * as ReactDom from 'react-dom';
import { sp } from '@pnp/sp';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'HelloWorldApplicationCustomizerStrings';
import GlobalNav from "./Components/GlobalNav";
import { IGlobalNavProps } from "./Components/IGlobalNavProps";
const LOG_SOURCE: string = 'HelloWorldApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHelloWorldApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
  termGroupId: string;
  termSetId: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HelloWorldApplicationCustomizer
  extends BaseApplicationCustomizer<IHelloWorldApplicationCustomizerProperties> {
  private _topPlaceholder: PlaceholderContent | undefined;
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    super.onInit().then(_ => {
      console.log("super oninit called");
      sp.setup({
        spfxContext: this.context
      });
    }).then((_) => {
      console.log("Sp Initilize");
      this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    });
    return Promise.resolve();

  }

  private _renderPlaceHolders(): void {
    console.log("HelloWorldApplicationCustomizer._renderPlaceHolders()");
    console.log(
      "Available placeholders: ",
      this.context.placeholderProvider.placeholderNames
        .map(name => PlaceholderName[name])
        .join(", ")
    );

    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error("The expected placeholder (Top) was not found.");
        return;
      }
      if (this.properties) {
        // Add refrence of react component to this file.
        const element: React.ReactElement<IGlobalNavProps> = React.createElement(
          GlobalNav,
          {
            termGroupId: this.properties.termGroupId,
            termSetId: this.properties.termSetId
          }
        );
        ReactDom.render(element, this._topPlaceholder.domElement);
      }
    }
  }

  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }

}
