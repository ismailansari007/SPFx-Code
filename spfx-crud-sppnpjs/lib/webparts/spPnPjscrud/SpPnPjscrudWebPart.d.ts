import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface ISpPnPjscrudWebPartProps {
    description: string;
}
export default class SpPnPjscrudWebPart extends BaseClientSideWebPart<ISpPnPjscrudWebPartProps> {
    render(): void;
    protected onDispose(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=SpPnPjscrudWebPart.d.ts.map