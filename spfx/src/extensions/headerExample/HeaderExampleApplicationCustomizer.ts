import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderName
} from '@microsoft/sp-application-base';

import * as strings from 'HeaderExampleApplicationCustomizerStrings';

const LOG_SOURCE: string = 'HeaderExampleApplicationCustomizer';

// Import the library
import "../../../../dist/sp-banner.js";
declare const SPBanner: (props: {
  el?: HTMLDivElement;
  context?: any;
  webUrl?: string;
}) => void;

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHeaderExampleApplicationCustomizerProperties {
  // This is an example; replace with your own property
  webUrl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HeaderExampleApplicationCustomizer
  extends BaseApplicationCustomizer<IHeaderExampleApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Handle the possible changes of the existence of placeholders
    this.context.placeholderProvider.changedEvent.add(this, this.renderBanner);

    // Resolve the event
    return Promise.resolve();
  }

  // Renders the banner
  private _elBanner: HTMLDivElement | undefined;
  private renderBanner(): void {
    // Do nothing if the banner exists
    if (this._elBanner) { return; }

    // Create the element
    this._elBanner = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top)?.domElement;

    // Render the banner
    SPBanner({
      el: this._elBanner,
      context: this.context,
      webUrl: this.properties.webUrl
    });
  }
}
