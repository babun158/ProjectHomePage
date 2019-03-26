import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
export interface IProjectHomeTilesWebPartProps {
    description: string;
}
export default class ProjectHomeTilesWebPart extends BaseClientSideWebPart<IProjectHomeTilesWebPartProps> {
    render(): void;
    protected readonly dataVersion: Version;
    FetchItems(): Promise<void>;
    AddNewTile(): void;
    UpdateItem(): void;
    DeleteItem(): void;
    Validation(): boolean;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
