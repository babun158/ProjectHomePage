var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './ProjectQuickLinksWebPart.module.scss';
import * as strings from 'ProjectQuickLinksWebPartStrings';
var ProjectQuickLinksWebPart = /** @class */ (function (_super) {
    __extends(ProjectQuickLinksWebPart, _super);
    function ProjectQuickLinksWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    ProjectQuickLinksWebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n      <div class=\"" + styles.projectQuickLinks + "\">\n        <div class=\"" + styles.container + "\">\n          <div class=\"" + styles.row + "\">\n            <div class=\"" + styles.column + "\">\n              <span class=\"" + styles.title + "\">Welcome to SharePoint!</span>\n              <p class=\"" + styles.subTitle + "\">Customize SharePoint experiences using Web Parts.</p>\n              <p class=\"" + styles.description + "\">" + escape(this.properties.description) + "</p>\n              <a href=\"https://aka.ms/spfx\" class=\"" + styles.button + "\">\n                <span class=\"" + styles.label + "\">Learn more</span>\n              </a>\n            </div>\n          </div>\n        </div>\n      </div>";
    };
    Object.defineProperty(ProjectQuickLinksWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    ProjectQuickLinksWebPart.prototype.getPropertyPaneConfiguration = function () {
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
    };
    return ProjectQuickLinksWebPart;
}(BaseClientSideWebPart));
export default ProjectQuickLinksWebPart;
//# sourceMappingURL=ProjectQuickLinksWebPart.js.map