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
import pnp from 'sp-pnp-js';
import { Web } from 'sp-pnp-js';
import { sp } from "sp-pnp-js";
import 'jquery';
// ADD NEW ITEM
function addItems(listName, listColumns) {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, pnp.sp.web.lists.getByTitle(listName).items.add(listColumns)];
                case 1:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
// ADD NEW ITEM WITH DOCUMENT
function additemsattachment(listName, file, listColumns) {
    return __awaiter(this, void 0, void 0, function () {
        var result;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, pnp.sp.web.getFolderByServerRelativeUrl(listName).files.add(file.name, file, true)];
                case 1:
                    result = _a.sent();
                    result.file.listItemAllFields.get().then(function (listItemAllFields) {
                        pnp.sp.web.lists.getByTitle(listName).items.getById(listItemAllFields.Id).update(listColumns);
                    });
                    return [2 /*return*/];
            }
        });
    });
}
// ADD NEW ITEM WITH IMAGE
function additemsimage(listName, filename, file, listColumns) {
    return __awaiter(this, void 0, void 0, function () {
        var result;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, pnp.sp.web.getFolderByServerRelativeUrl(listName).files.add(filename, file, true)
                        .then(function (result) {
                        result.file.listItemAllFields.get().then(function (listItemAllFields) {
                            return pnp.sp.web.lists.getByTitle(listName).items.getById(listItemAllFields.Id).update(listColumns);
                        });
                    })];
                case 1:
                    result = _a.sent();
                    return [2 /*return*/, result];
            }
        });
    });
}
// READ ITEMS
function readItems(listName, listColumns, topCount, orderBy, filterKey, filterValue) {
    return __awaiter(this, void 0, void 0, function () {
        var matchColumns, resultData;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    matchColumns = formString(listColumns);
                    if (!(filterKey == undefined)) return [3 /*break*/, 2];
                    return [4 /*yield*/, pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).top(topCount).orderBy(orderBy, false).get()];
                case 1:
                    resultData = _a.sent();
                    return [3 /*break*/, 4];
                case 2: return [4 /*yield*/, pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).filter("" + filterKey + " eq '" + filterValue + "'").top(topCount).orderBy(orderBy, false).get()];
                case 3:
                    resultData = _a.sent();
                    _a.label = 4;
                case 4: return [2 /*return*/, (resultData)];
            }
        });
    });
}
// READ Single ITEMS with Lookup
function readItem(listName, listColumns, topCount, orderBy, filterKey, filterValue, Lookupvalue) {
    return __awaiter(this, void 0, void 0, function () {
        var matchColumns, resultData;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    matchColumns = formString(listColumns);
                    if (!(Lookupvalue != "")) return [3 /*break*/, 1];
                    return [2 /*return*/, pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).expand(Lookupvalue).filter("" + filterKey + " eq '" + filterValue + "'").top(topCount).orderBy(orderBy, false).get()];
                case 1:
                    if (!(filterKey == undefined)) return [3 /*break*/, 3];
                    return [4 /*yield*/, pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).top(topCount).orderBy(orderBy, false).get()];
                case 2:
                    resultData = _a.sent();
                    return [3 /*break*/, 5];
                case 3: return [4 /*yield*/, pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).filter("" + filterKey + " eq '" + filterValue + "'").top(topCount).orderBy(orderBy, false).get()];
                case 4:
                    resultData = _a.sent();
                    _a.label = 5;
                case 5: return [2 /*return*/, (resultData)];
            }
        });
    });
}
// UPDATE ITEM
function updateItem(listName, id, listColumns) {
    return __awaiter(this, void 0, void 0, function () {
        var result;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, pnp.sp.web.lists.getByTitle(listName).items.getById(id).update(listColumns)];
                case 1:
                    result = _a.sent();
                    return [2 /*return*/, (result)];
            }
        });
    });
}
// DELETE ITEM
function deleteItem(listName, itemID) {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, pnp.sp.web.lists.getByTitle(listName).items.getById(itemID).delete()];
                case 1: return [2 /*return*/, _a.sent()];
            }
        });
    });
}
// BATCH DELETE - NOT YET TESTED
// async function batchDelete(listName: string, selectedArray: number[]) {  
//   let batch = sp.web.createBatch();
//   var arrayLen = selectedArray.length;  
//   for (var i =0; i<arrayLen;i++){
//     //await sp.web.lists.getByTitle(listName).items.getById(selectedArray[i]).inBatch(batch).delete().then(r => {
//     await sp.web.lists.getByTitle(listName).items.getById(selectedArray[i]).delete().then(r => {
//       console.log("deleted");
//     });
//   }
//   batch.execute().then(() => 
//   location.reload());
//   }
// GET FOLDER ONLY
function GetFolder(listName) {
    return __awaiter(this, void 0, void 0, function () {
        var folderList;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, pnp.sp.web.folders.getByName(listName).folders.expand('ListItemAllFields').get()];
                case 1:
                    folderList = _a.sent();
                    return [2 /*return*/, folderList];
            }
        });
    });
}
// // REMOVE FOLDER FROM DOC LIB
// async function DeleteFolder(listName: string, folderName: string){
//   console.log('common');
//   let confirm= await pnp.sp.web.folders.getByName(listName).folders.getByName(folderName).delete();
//   console.log(confirm);
//   return confirm;
// }
function batchDelete(listName, selectedArray) {
    return __awaiter(this, void 0, void 0, function () {
        var batch, arrayLen, i;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    batch = sp.web.createBatch();
                    arrayLen = selectedArray.length;
                    i = 0;
                    _a.label = 1;
                case 1:
                    if (!(i < arrayLen)) return [3 /*break*/, 4];
                    return [4 /*yield*/, sp.web.lists.getByTitle(listName).items.getById(selectedArray[i]).delete().then(function (r) { })];
                case 2:
                    _a.sent();
                    _a.label = 3;
                case 3:
                    i++;
                    return [3 /*break*/, 1];
                case 4:
                    batch.execute().then(function () { return location.reload(); });
                    return [2 /*return*/];
            }
        });
    });
}
// REMOVE FOLDER FROM DOC LIB
function DeleteFolder(listName, folderName) {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, pnp.sp.web.folders.getByName(listName).folders.getByName(folderName).delete()];
                case 1: return [2 /*return*/, _a.sent()];
            }
        });
    });
}
// CHECK USER IN GROUP
function checkUserinGroup(Componentname, email, callback) {
    return __awaiter(this, void 0, void 0, function () {
        var myitems;
        return __generator(this, function (_a) {
            pnp.sp.web.siteUsers
                .getByEmail(email)
                .groups.get()
                .then(function (items) {
                var currentComponent = Componentname;
                myitems = $.grep(items, function (obj, index) {
                    if (obj.Title.indexOf(currentComponent) != -1) {
                        return true;
                    }
                });
                callback(myitems.length);
            });
            return [2 /*return*/];
        });
    });
}
// GET ALL SUBSITES
function getListOfSubSites(url) {
    return __awaiter(this, void 0, void 0, function () {
        var result, my_web;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    my_web = new Web(url);
                    return [4 /*yield*/, my_web.webs.select().get()];
                case 1:
                    // let batch = web.createBatch();
                    result = _a.sent();
                    return [2 /*return*/, result];
            }
        });
    });
}
// GET LIST OF DOC LIBS
function getListOfDocLib(topCount, orderBy) {
    return __awaiter(this, void 0, void 0, function () {
        var result;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, pnp.sp.web.lists.filter('BaseTemplate eq 101').top(topCount).orderBy(orderBy).get()];
                case 1:
                    result = _a.sent();
                    return [2 /*return*/, result];
            }
        });
    });
}
// FORM STRING
function formString(listColumns) {
    var formattedString = "";
    for (var i = 0; i < listColumns.length; i++) {
        formattedString += listColumns[i] + ',';
    }
    return formattedString.slice(0, -1);
}
// FORMAT DATE
function formatDate(dateVal) {
    var date = new Date(dateVal);
    var year = date.getFullYear();
    var locale = "en-us";
    var month = date.toLocaleString(locale, { month: "long" });
    var dt = date.getDate();
    var dateString;
    if (dt < 10) {
        dateString = "0" + dt;
    }
    else
        dateString = dt.toString();
    return dateString + ' ' + month.substr(0, 3) + ',' + year;
}
function GetQueryStringParams(sParam) {
    var sPageURL = window.location.search.substring(1);
    var sURLVariables = sPageURL.split('&');
    for (var i = 0; i < sURLVariables.length; i++) {
        var sParameterName = sURLVariables[i].split('=');
        if (sParameterName[0] == sParam) {
            return sParameterName[1];
        }
    }
}
function base64ToArrayBuffer(base64) {
    var binary_string = window.atob(base64);
    var len = binary_string.length;
    var bytes = new Uint8Array(len);
    for (var i = 0; i < len; i++) {
        bytes[i] = binary_string.charCodeAt(i);
    }
    return bytes.buffer;
}
export { getListOfDocLib, getListOfSubSites, addItems, readItems, readItem, additemsimage, deleteItem, updateItem, DeleteFolder, GetFolder, formString, additemsattachment, checkUserinGroup, batchDelete, formatDate, GetQueryStringParams, base64ToArrayBuffer };
//# sourceMappingURL=commonJS.js.map