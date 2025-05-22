import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    type IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import * as strings from 'UserInfoWebpartWebPartStrings';
import UserInfoWebpart from './components/UserInfoWebpart';
import { IUserInfoWebpartProps } from './components/IUserInfoWebpartProps';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base';


export interface IUserInfoWebpartWebPartProps {
    description: string;
    fontSize: string;
}

export default class UserInfoWebpartWebPart extends BaseClientSideWebPart<IUserInfoWebpartWebPartProps> {

    private _isDarkTheme: boolean = false;
    private _environmentMessage: string = '';
    private themeProvider: ThemeProvider;
    private themeVariant: IReadonlyTheme | undefined;
    private _siteTitleFontName: string;

    public render(): void {
        try {
            const element: React.ReactElement<IUserInfoWebpartProps> = React.createElement(
                UserInfoWebpart,
                {
                    description: this.properties.description,
                    isDarkTheme: this._isDarkTheme,
                    environmentMessage: this._environmentMessage,
                    hasTeamsContext: !!this.context.sdks.microsoftTeams,
                    userDisplayName: this.context.pageContext.user.displayName,
                    fontSize: this.properties.fontSize,
                    siteTitleFontName: this._siteTitleFontName
                }
            );

            ReactDom.render(element, this.domElement);
        }
        catch (error) {
            console.error('Błąd w renderowaniu webparta:', error);
        }
    }

    protected onInit(): Promise<void> {
        this.themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

        this.themeProvider.themeChangedEvent.add(this, this.handleThemeChanged);

        this.themeVariant = this.themeProvider.tryGetTheme();

        if (this.themeVariant) { 
            // @ts-expect-error // fontSlots is not in IReadonlyTheme type definition, but exists in runtime
            this._siteTitleFontName = this.themeVariant?.fontSlots?.title?.fontFamily || "Domyślna czcionka tytułu";
        }

        return super.onInit();
    }

    private handleThemeChanged(args: ThemeChangedEventArgs): void {
        this.themeVariant = args.theme;

        // @ts-expect-error // fontSlots is not in IReadonlyTheme type definition, but exists in runtime
        this._siteTitleFontName = this.themeVariant?.fontSlots?.title?.fontFamily || "Domyślna czcionka tytułu";

        this.render(); 
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
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
                                }),
                                PropertyPaneTextField('fontSize', {
                                    label: "Rozmiar czcionki (np. '16px', '20px', '1.5rem')",
                                    value: this.properties.fontSize || "16px"
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}