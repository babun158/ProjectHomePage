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
import * as strings from 'ProjectHomeTilesWebPartStrings';
import * as $ from "jquery";
import { addItems, readItems, deleteItem, updateItem } from '../../commonJS';
var ProjectHomeTilesWebPart = /** @class */ (function (_super) {
    __extends(ProjectHomeTilesWebPart, _super);
    function ProjectHomeTilesWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    ProjectHomeTilesWebPart.prototype.render = function () {
        var _this = this;
        this.domElement.innerHTML = "\n    <section class=\"gallery-section\">\n    </section>\n    ";
        var ModalHTML = '<div class="modal fade" id="Addmodal" tabindex="-1" role="dialog" aria-labelledby="basicModal" aria-hidden="true">' +
            '<div class="modal-dialog modal-md">' +
            '<div class="modal-content">' +
            '<div class="modal-header">' +
            '<h4 class="modal-title" id="myModalLabel">Add New Favourites</h4>' +
            '<button type="button" class="close" data-dismiss="modal" aria-label="Close"> <span class="icon-remove"></span> </button>' +
            '</div>' +
            '<div class="modal-body">' +
            '<div class="col-xs-12 form-element">' +
            '<label class="required">Title</label>' +
            '<input type="text" id="txtTitle" placeholder="Title of the site or link" class="form-control">' +
            '</div>' +
            '<div class="col-xs-12 form-element">' +
            '<label class="required">URL</label>' +
            '<input type="text" id="txtLinkUrl" class="form-control">' +
            '<span>Please enter the Link URL in the following format : https://www.bloomholding.com</span>' +
            '<input type="text" id="txtID" style="display:none" class="form-control">' +
            '</div>' +
            '</div>' +
            '<div class="modal-footer">' +
            '<div class="col-xs-12 form-element"> <a id="btnAddSubmit" href="#" class="s-button">Submit</a> <a id="btnEditSubmit" href="#" class="s-button">Submit</a><label id="lblwait" style="display:none;float:left;">Please Wait...</label>  <a href="#" id="btnDelete" class="r-btn"><i class="icon-delete"></i> Delete</a></div>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '</div>';
        $("body").after(ModalHTML);
        var siteurl = this.context.pageContext.site.absoluteUrl;
        this.FetchItems();
        var Addevent = $('#btnAddSubmit');
        Addevent.on("click", function (e) { return _this.AddNewTile(); });
        var EditEvent = $('#btnEditSubmit');
        EditEvent.on("click", function (e) { return _this.UpdateItem(); });
        var DeleteEvent = $('#btnDelete');
        DeleteEvent.on("click", function (e) { return _this.DeleteItem(); });
    };
    Object.defineProperty(ProjectHomeTilesWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    ProjectHomeTilesWebPart.prototype.FetchItems = function () {
        return __awaiter(this, void 0, void 0, function () {
            var listName, columnArray, Username, GetListItems, HTML, Tiles, FavIcon, i;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        listName = "Tiles";
                        columnArray = ["Username", "LinkURL", "Title", "ID"];
                        Username = this.context.pageContext.user.displayName;
                        return [4 /*yield*/, readItems(listName, columnArray, 10, "ID", "Username", Username)];
                    case 1:
                        GetListItems = _a.sent();
                        HTML = "";
                        Tiles = [];
                        if (GetListItems.length > 0) {
                            for (i = 0; i < GetListItems.length; i++) {
                                if (GetListItems[i].LinkURL.Url.toLowerCase().indexOf("bloom") != -1) {
                                    FavIcon = this.context.pageContext.site.absoluteUrl + "/_catalogs/masterpage/BloomProject/images/favicon.ico";
                                }
                                else {
                                    FavIcon = GetListItems[i].LinkURL.Url + '/favicon.ico';
                                }
                                if (i == 0) {
                                    HTML += '<div class="col-sm-3 col-xs-12  pad-left0">' +
                                        '<div class="img-gallery"><div class="exp-img"><div align="center" class="small-icon"> <img src="' + FavIcon + '"></div></div> <a target="_blank" class="head-h3" href="' + GetListItems[i].LinkURL.Url + '" id="' + GetListItems[i].ID + '" >' + GetListItems[i].Title + '</a> <a class="icon-more TileEdit" style="cursor: pointer;" id="' + GetListItems[i].ID + '" data-toggle="modal" data-target="#Addmodal"></a></div>' +
                                        //'<div class="img-gallery"> <img src="images/announce-listimg1.jpg"> <a href="#" class="TileEdit" id="'+GetListItems[i].ID+'" >'+GetListItems[i].Title+'<i class="icon-more" data-toggle="modal" data-target="#Addmodal"></i></a> </div>'+
                                        '</div>';
                                }
                                else if (i == 3) {
                                    HTML += '<div class="col-sm-3 col-xs-12 over-right">' +
                                        '<div class="img-gallery"><div class="exp-img"><div align="center" class="small-icon"> <img src="' + FavIcon + '"></div></div> <a target="_blank" class="head-h3" href="' + GetListItems[i].LinkURL.Url + '" id="' + GetListItems[i].ID + '" >' + GetListItems[i].Title + '</a> <a class="icon-more TileEdit" style="cursor: pointer;" id="' + GetListItems[i].ID + '" data-toggle="modal" data-target="#Addmodal"></a></div>' +
                                        //'<div class="img-gallery"> <img src="images/announce-listimg1.jpg"> <a href="#" class="TileEdit" id="'+GetListItems[i].ID+'" >'+GetListItems[i].Title+'<i class="icon-more" data-toggle="modal" data-target="#Addmodal"></i></a> </div>'+
                                        '</div>';
                                }
                                else {
                                    HTML += '<div class="col-sm-3 col-xs-12">' +
                                        '<div class="img-gallery"><div class="exp-img"><div align="center" class="small-icon"> <img src="' + FavIcon + '"></div></div> <a target="_blank" class="head-h3" href="' + GetListItems[i].LinkURL.Url + '" id="' + GetListItems[i].ID + '" >' + GetListItems[i].Title + '</a> <a class="icon-more TileEdit" style="cursor: pointer;" id="' + GetListItems[i].ID + '" data-toggle="modal" data-target="#Addmodal"></a></div>' +
                                        //'<div class="img-gallery"> <img src="images/announce-listimg1.jpg"> <a href="#" class="TileEdit" id="'+GetListItems[i].ID+'" >'+GetListItems[i].Title+'<i class="icon-more" data-toggle="modal" data-target="#Addmodal"></i></a> </div>'+
                                        '</div>';
                                }
                            }
                        }
                        if (GetListItems.length < 4) {
                            HTML += '<div class="col-sm-3 col-xs-12 over-right">' +
                                '<div class="no-img"> <a href="#" class="AddTile" data-toggle="modal" data-target="#Addmodal"><i class="icon-add"></i>Add Favourites</a></div>' +
                                '</div>';
                        }
                        $('.gallery-section').append(HTML);
                        $('.TileEdit').click(function (event) {
                            $('#myModalLabel').text("EDIT FAVOURITES");
                            var TileID = $(this).attr("id");
                            $('#btnAddSubmit').hide();
                            $('#btnEditSubmit').show();
                            $('#btnDelete').show();
                            $('#txtTitle').val("");
                            $('#txtLinkUrl').val("");
                            var GetListItems = readItems(listName, columnArray, 10, "Modified", "ID", parseInt(TileID));
                            GetListItems.then(function (result) {
                                $('#txtID').val(result[0].ID);
                                $('#txtTitle').val(result[0].Title);
                                $('#txtLinkUrl').val(result[0].LinkURL.Url);
                            });
                        });
                        $('.icon-remove').click(function (event) {
                            $('#txtTitle').val("");
                            $('#txtLinkUrl').val("");
                            $('#btnAddSubmit').show();
                            $('#btnEditSubmit').hide();
                        });
                        $('.AddTile').click(function (event) {
                            $('#myModalLabel').text("ADD NEW FAVOURITES");
                            $('#txtTitle').val("");
                            $('#txtLinkUrl').val("");
                            $('#btnAddSubmit').show();
                            $('#btnEditSubmit').hide();
                            $('#Addmodal').show();
                            $('#btnDelete').hide();
                        });
                        return [2 /*return*/];
                }
            });
        });
    };
    ProjectHomeTilesWebPart.prototype.AddNewTile = function () {
        if ($('.ajs-message').length > 0) {
            $('.ajs-message').remove();
        }
        if (this.Validation()) {
            $('#btnAddSubmit').hide();
            $('#lblwait').show();
            var listName = "Tiles";
            var itemObj = {
                Title: $('#txtTitle').val(),
                Username: this.context.pageContext.user.displayName,
                LinkURL: {
                    "__metadata": {
                        "type": "SP.FieldUrlValue"
                    },
                    Url: $('#txtLinkUrl').val()
                }
            };
            addItems(listName, itemObj).then(function (result) {
                location.reload();
            });
        }
    };
    ProjectHomeTilesWebPart.prototype.UpdateItem = function () {
        if ($('.ajs-message').length > 0) {
            $('.ajs-message').remove();
        }
        if (this.Validation()) {
            $('#btnEditSubmit').hide();
            $('#lblwait').show();
            var listName = "Tiles";
            var columnArray = ["Title", "URL"];
            var itemId = +$('#txtID').val();
            var itemObj = {
                Title: $('#txtTitle').val(),
                Username: this.context.pageContext.user.displayName,
                LinkURL: {
                    "__metadata": {
                        "type": "SP.FieldUrlValue"
                    },
                    Url: $('#txtLinkUrl').val()
                }
            };
            updateItem(listName, itemId, itemObj).then(function (result) {
                location.reload();
            });
        }
    };
    ProjectHomeTilesWebPart.prototype.DeleteItem = function () {
        var strconfirm = "Are you sure you want to delete ?";
        var _this = this;
        alertify.confirm('Confirmation', strconfirm, function () {
            var listName = "Tiles";
            var itemId = +$('#txtID').val();
            deleteItem(listName, itemId).then(function (result) {
                location.reload();
            });
        }, function () {
        }).set('closable', false);
    };
    ProjectHomeTilesWebPart.prototype.Validation = function () {
        var regexp = /^(https?|s?ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i;
        if (!$('#txtTitle').val()) {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error("Please Enter the Title");
            return false;
        }
        else if (!$('#txtLinkUrl').val()) {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error("Please Enter the URL");
            return false;
        }
        else if (!regexp.test($('#txtLinkUrl').val().toString().trim())) {
            //$('#txtHyper').focus();
            alertify.set('notifier', 'position', 'top-right');
            alertify.error("Please Enter Link URL Correctly");
            return false;
        }
        return true;
    };
    ProjectHomeTilesWebPart.prototype.getPropertyPaneConfiguration = function () {
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
    return ProjectHomeTilesWebPart;
}(BaseClientSideWebPart));
export default ProjectHomeTilesWebPart;
//# sourceMappingURL=ProjectHomeTilesWebPart.js.map