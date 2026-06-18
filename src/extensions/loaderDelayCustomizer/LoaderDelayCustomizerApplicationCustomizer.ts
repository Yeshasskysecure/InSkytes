import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'LoaderDelayCustomizerApplicationCustomizerStrings';

const LOG_SOURCE: string = 'LoaderDelayCustomizerApplicationCustomizer';

export interface ILoaderDelayCustomizerApplicationCustomizerProperties {
  testMessage: string;
}

export default class LoaderDelayCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<ILoaderDelayCustomizerApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    if (!document.getElementById('spfx-loader-delay')) {
      const style = document.createElement('style');
      style.id = 'spfx-loader-delay';
      style.textContent = `
        [id^="embeddedPageShell"] {
          opacity: 0 !important;
          animation: spLoaderFadeIn 0s 0.35s forwards !important;
          transition: none !important;
        }
        @keyframes spLoaderFadeIn {
          to { opacity: 1; }
        }
      `;
      document.head.appendChild(style);
    }

    return Promise.resolve();
  }
}