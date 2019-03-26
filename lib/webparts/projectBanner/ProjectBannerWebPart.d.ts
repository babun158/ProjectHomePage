import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import 'jquery';
export interface IProjectBannerWebPartProps {
    description: string;
}
export default class ProjectBannerWebPart extends BaseClientSideWebPart<IProjectBannerWebPartProps> {
    userflag: boolean;
    render(): void;
    viewlistitemdesign(): void;
    viewpageRedirect(siteURL: any): void;
    BannerPage(userflag: any): Promise<void>;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
