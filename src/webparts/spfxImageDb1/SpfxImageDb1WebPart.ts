import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ThemeProvider } from '@microsoft/sp-component-base';
import * as strings from 'SpfxImageDb1WebPartStrings';
import SpfxImageDb1 from './components/SpfxImageDb1';
import { ISpfxImageDb1Props } from './components/ISpfxImageDb1Props';
import {
  IdPrefixProvider,
  FluentProvider,
  FluentProviderProps,
  webLightTheme,
  webDarkTheme,
  Theme,
} from "@fluentui/react-components";
import { createV9Theme } from "@fluentui/react-migration-v8-v9";

export interface ISpfxImageDb1WebPartProps {
  description: string;
}

export default class SpfxImageDb1WebPart extends BaseClientSideWebPart<ISpfxImageDb1WebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _theme: Theme = webLightTheme;
  private _themeProvider: ThemeProvider | undefined;

  public render(): void {
    const element: React.ReactElement<ISpfxImageDb1Props> = React.createElement(
      SpfxImageDb1,
      {
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    const fluentElement: React.ReactElement<FluentProviderProps> = React.createElement(
      IdPrefixProvider,
      { value: "spfx-image-db1" },
      React.createElement(
        FluentProvider,
        {
          theme: this._theme,
          className: "spfx-image-db1",
        },
        element
      )
    );

    ReactDom.render(fluentElement, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    const theme: IReadonlyTheme | undefined = this._themeProvider.tryGetTheme();
    if (theme) {
      this.applyThemeSettings(theme);
    }

    this._themeProvider.themeChangedEvent.add(this, this.onThemeChanged.bind(this));

    return this._getEnvironmentMessage().then((message: string) => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  private applyThemeSettings(theme: IReadonlyTheme | undefined): void {
    if (!theme) return;

    this._isDarkTheme = !!theme.isInverted;

    this._theme = createV9Theme(
      theme as any,
      this._isDarkTheme ? webDarkTheme : webLightTheme
    );

    this.render();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this.applyThemeSettings(currentTheme);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
