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
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'ProjectBannerWebPartStrings';
import 'jquery';
import { readItems, checkUserinGroup } from '../../commonJS';
var ProjectBannerWebPart = /** @class */ (function (_super) {
    __extends(ProjectBannerWebPart, _super);
    function ProjectBannerWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.userflag = false;
        return _this;
    }
    ProjectBannerWebPart.prototype.render = function () {
        var _this = this;
        //Checking user details in group
        checkUserinGroup("Banners", this.context.pageContext.user.email, function (result) {
            if (result == 1) {
                _this.userflag = true;
            }
            _this.viewlistitemdesign();
        });
    };
    ProjectBannerWebPart.prototype.viewlistitemdesign = function () {
        var _this = this;
        var siteURL = this.context.pageContext.web.absoluteUrl;
        this.domElement.innerHTML =
            '<section class="banner-section">' +
                "<div class='ban-section'>" +
                '<h3 class="tt-head">' + this.context.pageContext.web.title + '</h3>' +
                '<div id="carousel-banner" class="carousel carousel-fade" data-ride="carousel">' +
                '<ol class="carousel-indicators banner-carousel">' +
                '</ol>' +
                '</div>' +
                '<div id="addEvents" class="event-add" style="Display:none">' +
                //'<h3 class="banner-itemview" style="cursor:pointer">UPDATE COVERAGE EVENTS <a href="' + siteURL + '/Pages/ListView.aspx?CName=Banners"></a><i href="' + siteURL + '/Pages/AddListItem.aspx?CName=Banners" class="icon-add"></i></h3>' +
                '<h3 class="banner-itemview" style="cursor:pointer">VIEW COVERAGE EVENTS <a href="' + siteURL + '/Pages/AddListItem.aspx?CName=Banners"><i class="icon-add"></i></a> </h3>' +
                '</div>' +
                "</div>" +
                '<section>';
        ;
        this.BannerPage(this.userflag);
        var viewevent = document.getElementsByClassName('banner-itemview');
        for (var i = 0; i < viewevent.length; i++) {
            viewevent[i].addEventListener("click", function (e) { return _this.viewpageRedirect(siteURL); });
        }
        $('#carousel-banner').carousel({ interval: 3000 });
    };
    ProjectBannerWebPart.prototype.viewpageRedirect = function (siteURL) {
        window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=Banners";
    };
    ProjectBannerWebPart.prototype.BannerPage = function (userflag) {
        return __awaiter(this, void 0, void 0, function () {
            var renderhtml, renderliitems, count, activeflag, bannerItems, siteURL, i, DottedTitle;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        renderhtml = '<div class="carousel-inner" role="listbox">';
                        renderliitems = "";
                        siteURL = this.context.pageContext.site.absoluteUrl;
                        return [4 /*yield*/, readItems("Banners", ["Title", "Modified", "LinkURL", "Display", "BannerContent", "Image"], 3, "Modified", "Display", 1)];
                    case 1:
                        bannerItems = _a.sent();
                        if (bannerItems.length > 0) {
                            for (i = 0; i < bannerItems.length; i++) {
                                if (i == 0) {
                                    activeflag = "active";
                                }
                                else {
                                    activeflag = "";
                                }
                                renderliitems += '<li data-slide-to="' + i + '" data-target="#carousel-banner" class="' + activeflag + '">' + '</li>';
                                renderhtml += '<div class="item ' + activeflag + '">';
                                renderhtml += '<img src="' + bannerItems[i].Image.Url + '" style="max-height: 319px;"alt="Slide" title="Slide" />';
                                renderhtml += '<div class="carousel-caption">';
                                DottedTitle = bannerItems[i].Title;
                                if (DottedTitle.length > 65) {
                                    DottedTitle = DottedTitle.substring(0, 65) + "...";
                                }
                                if (bannerItems[i].LinkURL !== null) {
                                    renderhtml += '<div align="center">' + '<a href="' + bannerItems[i].LinkURL.Url + '" class="wow fadeInRight" style="visibility: visible; animation-name: fadeInRight;">lEARN mORE</a>' + '</div>';
                                }
                                renderhtml += '</div>';
                                renderhtml += '</div>';
                            }
                        }
                        else if (bannerItems.length == 0) {
                            activeflag = "active";
                            renderliitems += '<li data-slide-to="1" data-target="#carousel-banner" class="' + activeflag + '"></li>';
                            renderhtml += '<div class="item ' + activeflag + '">';
                            renderhtml += '<img src="../../../../_catalogs/masterpage/BloomHomepage/images/logo.png" style="max-height: 319px;"alt="Slide" title="Slide" />';
                            renderhtml += '<div class="carousel-caption">';
                            renderhtml += '<p></p>';
                            renderhtml += '<h3 class="wow fadeInRight no-data" style="visibility: visible; animation-name: fadeInRight;">No Banner Image To Display</h3>';
                            renderhtml += '</div>';
                            renderhtml += '</div>';
                        }
                        renderhtml += '</div>';
                        renderhtml += '<!-- Left and right controls -->';
                        renderhtml += '<a class="left carousel-control" href="#carousel-banner" data-slide="prev">';
                        renderhtml += '<span class="glyphicon glyphicon-chevron-left"></span>';
                        renderhtml += '<span class="sr-only">Previous</span>';
                        renderhtml += '</a>';
                        renderhtml += '<a class="right carousel-control" href="#carousel-banner" data-slide="next">';
                        renderhtml += '<span class="glyphicon glyphicon-chevron-right"></span>';
                        renderhtml += '<span class="sr-only">Next</span>';
                        renderhtml += '</a>';
                        //renderhtml += '<div id="addEvents" class="event-add">'+'<h3>UPDATE COVERAGE EVENTS <a href="#"><i class="icon-add"></i></a> </h3>'+'</div>';
                        if (userflag == false) {
                            $('#addEvents').hide();
                        }
                        else {
                            $('#addEvents').show();
                        }
                        $(".banner-carousel").append(renderliitems);
                        $(".banner-carousel").after(renderhtml);
                        return [2 /*return*/];
                }
            });
        });
    };
    Object.defineProperty(ProjectBannerWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    ProjectBannerWebPart.prototype.getPropertyPaneConfiguration = function () {
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
    return ProjectBannerWebPart;
}(BaseClientSideWebPart));
export default ProjectBannerWebPart;
//# sourceMappingURL=ProjectBannerWebPart.js.map