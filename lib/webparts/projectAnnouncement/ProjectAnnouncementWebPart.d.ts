import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
export interface IProjectAnnouncementWebPartProps {
    description: string;
}
export default class ProjectAnnouncementWebPart extends BaseClientSideWebPart<IProjectAnnouncementWebPartProps> {
    userflag: boolean;
    render(): void;
    getAnnouncements(userflag: any): Promise<void>;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
