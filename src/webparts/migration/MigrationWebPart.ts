import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MigrationWebPartStrings';
import Migration from './components/Migration';
import { IMigrationProps } from './components/IMigrationProps';

export interface IMigrationWebPartProps {
  description: string;
}

export default class MigrationWebPart extends BaseClientSideWebPart<IMigrationWebPartProps> {
  private static readonly earlyHiderStyleId = 'ikn-early-hider';
  private _shellObserver: MutationObserver | null = null;
  private _shellSuppressionFallbackId: number | undefined;

  public async onInit(): Promise<void> {
    const normalized = new URL(window.location.href);
    const hasWebViewListEnv = (normalized.searchParams.get('env') || '').toLowerCase() === 'webviewlist';

    if (!hasWebViewListEnv) {
      normalized.searchParams.set('env', 'WebViewList');
      window.location.replace(normalized.toString());
      return;
    }

    // MUST BE FIRST - before super.onInit(), before anything.
    if (!window.history.state?.sameTabNav) {
      document.documentElement.style.cssText =
        'visibility:hidden!important;background:#fff!important';

      const s = document.createElement('style');
      s.id = 'ikn-early-hider';
      s.textContent = 'html{visibility:hidden!important;background:#fff!important}';
      const earlyStyleHost = document.head || document.documentElement;
      earlyStyleHost.insertBefore(s, earlyStyleHost.firstChild);
    }

    await super.onInit();
    this._startShellSuppression();
  }

  public render(): void {

    // Show skeleton immediately
    this.domElement.innerHTML = `
      <div id="ik-skeleton-wrapper" style="padding: 20px;">
        <style>
          @keyframes ikShimmer {
            0% { background-position: 200% 0; }
            100% { background-position: -200% 0; }
          }
          .ik-skel {
            background: linear-gradient(90deg, #f0f0f0 25%, #e8e8e8 50%, #f0f0f0 75%);
            background-size: 200% 100%;
            animation: ikShimmer 1.5s infinite;
            border-radius: 4px;
          }
        </style>

        <!-- Header row -->
        <div style="display:flex; align-items:center; gap:12px; margin-bottom:24px;">
          <div class="ik-skel" style="height:32px; width:200px;"></div>
          <div style="flex:1;"></div>
          <div class="ik-skel" style="height:32px; width:120px;"></div>
          <div class="ik-skel" style="height:32px; width:120px;"></div>
        </div>

        <!-- Count line -->
        <div class="ik-skel" style="height:16px; width:280px; margin-bottom:20px;"></div>

        <!-- Document rows -->
        ${[1,2,3,4,5,6].map(() => `
          <div style="display:flex; gap:16px; padding:16px 0; border-bottom:1px solid #f0f0f0; align-items:center;">
            <div style="flex:1;">
              <div class="ik-skel" style="height:18px; width:60%; margin-bottom:8px;"></div>
              <div class="ik-skel" style="height:14px; width:85%; margin-bottom:6px;"></div>
              <div class="ik-skel" style="height:14px; width:40%;"></div>
            </div>
            <div style="display:flex; flex-direction:column; gap:8px; align-items:flex-end; flex-shrink:0;">
              <div class="ik-skel" style="width:80px; height:14px;"></div>
              <div class="ik-skel" style="width:60px; height:14px;"></div>
            </div>
            <div style="display:flex; gap:8px; flex-shrink:0;">
              <div class="ik-skel" style="width:64px; height:32px; border-radius:4px;"></div>
              <div class="ik-skel" style="width:80px; height:32px; border-radius:4px;"></div>
            </div>
          </div>
        `).join('')}
      </div>
    `;

    // Wait for skeleton to paint before mounting React
    // requestAnimationFrame ensures browser has painted the skeleton first
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        const urlParams = new URLSearchParams(window.location.search);
        const projectId = urlParams.get('projectId') || undefined;
        const { detectCurrentPage, parseUrlParams } =
          require('./services/permalinkService');

        const currentPage = detectCurrentPage();
        const parsedParams = parseUrlParams();

        const element: React.ReactElement<IMigrationProps> = React.createElement(
          Migration,
          {
            description: this.properties.description,
            context: this.context,
            projectId: projectId,
            currentPage: currentPage,
            initialDocumentId: parsedParams.docId,
            initialSearchQuery: parsedParams.query
          }
        );

        ReactDom.render(element, this.domElement);
      });
    });
  }

  protected onDispose(): void {
    this._stopShellSuppression();
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  private _startShellSuppression(): void {
    if (typeof window === 'undefined' || typeof document === 'undefined' || !document.body) {
      return;
    }

    if (window.history.state?.sameTabNav) {
      return;
    }

    const wpElement = this.domElement;
    wpElement.style.opacity = '0';

    document.documentElement.style.background = '#fff';
    document.body.style.background = '#fff';
    document.body.style.overflow = 'hidden';

    const ancestors = new Set<Element>();
    let el: Element | null = wpElement;

    while (el) {
      ancestors.add(el);
      el = el.parentElement;
    }

    const isAppPortalElement = (element: HTMLElement): boolean => {
      return String(element.className || '').indexOf('chatPortalRoot') !== -1;
    };

    const hideElement = (elementToHide: HTMLElement): void => {
      if (elementToHide.dataset.iknShellHidden === 'true') {
        return;
      }

      elementToHide.dataset.iknShellHidden = 'true';
      elementToHide.style.setProperty('display', 'none', 'important');
      elementToHide.style.setProperty('visibility', 'hidden', 'important');
      elementToHide.style.setProperty('pointer-events', 'none', 'important');
    };

    const hideNonAncestors = (): void => {
      document.body.childNodes.forEach((node) => {
        if (node.nodeType !== 1) {
          return;
        }

        const bodyChild = node as HTMLElement;
        if (!ancestors.has(bodyChild) && !isAppPortalElement(bodyChild)) {
          hideElement(bodyChild);
        }
      });

      ancestors.forEach((ancestor) => {
        if (ancestor === document.body || ancestor === document.documentElement) {
          return;
        }

        Array.from(ancestor.parentElement?.children || []).forEach((sibling) => {
          const siblingElement = sibling as HTMLElement;
          if (!ancestors.has(sibling) && !isAppPortalElement(siblingElement)) {
            hideElement(siblingElement);
          }
        });
      });
    };

    hideNonAncestors();

    this._shellObserver = new MutationObserver(() => {
      hideNonAncestors();
    });

    this._shellObserver.observe(document.body, {
      childList: true,
      subtree: false
    });

    ancestors.forEach((ancestor) => {
      if (ancestor === document.body || ancestor === document.documentElement) {
        return;
      }

      const ancestorElement = ancestor as HTMLElement;
      ancestorElement.style.setProperty('width', '100vw', 'important');
      ancestorElement.style.setProperty('max-width', '100vw', 'important');
      ancestorElement.style.setProperty('padding', '0', 'important');
      ancestorElement.style.setProperty('margin', '0', 'important');
      ancestorElement.style.setProperty('display', 'block', 'important');
    });

    const previousReady = (window as any).__appReady;
    (window as any).__appReady = (): void => {
      if (typeof previousReady === 'function') {
        previousReady();
      }
      this._stopShellSuppression();
      wpElement.style.transition = 'opacity 250ms ease-out';
      window.requestAnimationFrame(() => {
        wpElement.style.opacity = '1';
      });
    };

    this._shellSuppressionFallbackId = window.setTimeout(() => {
      document.documentElement.style.visibility = '';
      document.documentElement.style.background = '';
      const early = document.getElementById('ikn-early-hider');
      if (early) early.remove();
      this._stopShellSuppression();
      wpElement.style.opacity = '1';
    }, 5000);
  }

  private _stopShellSuppression(): void {
    if (this._shellSuppressionFallbackId !== undefined) {
      window.clearTimeout(this._shellSuppressionFallbackId);
      this._shellSuppressionFallbackId = undefined;
    }

    if (this._shellObserver) {
      this._shellObserver.disconnect();
      this._shellObserver = null;
    }

    if (typeof document !== 'undefined' && document.body) {
      document.documentElement.style.visibility = '';
      document.documentElement.style.background = '';
      document.body.style.overflow = '';
      document.getElementById(MigrationWebPart.earlyHiderStyleId)?.remove();
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
