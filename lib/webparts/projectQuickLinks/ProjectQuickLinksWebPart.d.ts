import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
export interface IProjectQuickLinksWebPartProps {
    description: string;
}
export default class ProjectQuickLinksWebPart extends BaseClientSideWebPart<IProjectQuickLinksWebPartProps> {
    render(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
