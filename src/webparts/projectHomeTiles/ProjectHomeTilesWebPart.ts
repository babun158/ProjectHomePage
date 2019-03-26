import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ProjectHomeTilesWebPart.module.scss';
import * as strings from 'ProjectHomeTilesWebPartStrings';
import * as $ from "jquery";
import { addItems, readItems, deleteItem, updateItem  } from '../../commonJS';
declare var alertify:any;
export interface IProjectHomeTilesWebPartProps {
  description: string;
}

export default class ProjectHomeTilesWebPart extends BaseClientSideWebPart<IProjectHomeTilesWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <section class="gallery-section">
    </section>
    `;

  var ModalHTML = '<div class="modal fade" id="Addmodal" tabindex="-1" role="dialog" aria-labelledby="basicModal" aria-hidden="true">'+
  '<div class="modal-dialog modal-md">'+
    '<div class="modal-content">'+
      '<div class="modal-header">'+
        '<h4 class="modal-title" id="myModalLabel">Add New Favourites</h4>'+
        '<button type="button" class="close" data-dismiss="modal" aria-label="Close"> <span class="icon-remove"></span> </button>'+
      '</div>'+
      '<div class="modal-body">'+
        '<div class="col-xs-12 form-element">'+
          '<label class="required">Title</label>'+
          '<input type="text" id="txtTitle" placeholder="Title of the site or link" class="form-control">'+
        '</div>'+
        '<div class="col-xs-12 form-element">'+
          '<label class="required">URL</label>'+
          '<input type="text" id="txtLinkUrl" class="form-control">'+
          '<span>Please enter the Link URL in the following format : https://www.bloomholding.com</span>'+
          '<input type="text" id="txtID" style="display:none" class="form-control">'+
        '</div>'+
      '</div>'+
      '<div class="modal-footer">'+
        '<div class="col-xs-12 form-element"> <a id="btnAddSubmit" href="#" class="s-button">Submit</a> <a id="btnEditSubmit" href="#" class="s-button">Submit</a><label id="lblwait" style="display:none;float:left;">Please Wait...</label>  <a href="#" id="btnDelete" class="r-btn"><i class="icon-delete"></i> Delete</a></div>'+
      '</div>'+
    '</div>'+
  '</div>'+
'</div>';

      $("body").after(ModalHTML);
      var siteurl = this.context.pageContext.site.absoluteUrl;
      this.FetchItems();

      let Addevent = $('#btnAddSubmit');
      Addevent.on("click", (e: Event) => this.AddNewTile());

      let EditEvent = $('#btnEditSubmit');
      EditEvent.on("click", (e: Event) => this.UpdateItem());

      let DeleteEvent = $('#btnDelete');
      DeleteEvent.on("click", (e: Event) => this.DeleteItem());
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  async FetchItems(){
    var listName = "Tiles";
    let columnArray = ["Username","LinkURL","Title","ID"];
    var Username = this.context.pageContext.user.displayName;
    let GetListItems = await readItems(listName, columnArray, 10, "ID","Username",Username);
    var HTML = "";
    var Tiles = [];
    var FavIcon;
    if(GetListItems.length > 0)
    {
      for(var i=0; i<GetListItems.length; i++)
      {
        if(GetListItems[i].LinkURL.Url.toLowerCase().indexOf("bloom") != -1)
        {
          FavIcon = this.context.pageContext.site.absoluteUrl + "/_catalogs/masterpage/BloomProject/images/favicon.ico";
        }
        else{
          FavIcon = GetListItems[i].LinkURL.Url+'/favicon.ico';
        }
        if(i==0)
        {
          HTML += '<div class="col-sm-3 col-xs-12  pad-left0">'+
                    '<div class="img-gallery"><div class="exp-img"><div align="center" class="small-icon"> <img src="'+FavIcon+'"></div></div> <a target="_blank" class="head-h3" href="'+GetListItems[i].LinkURL.Url+'" id="'+GetListItems[i].ID+'" >'+GetListItems[i].Title+'</a> <a class="icon-more TileEdit" style="cursor: pointer;" id="'+GetListItems[i].ID+'" data-toggle="modal" data-target="#Addmodal"></a></div>'+
                      //'<div class="img-gallery"> <img src="images/announce-listimg1.jpg"> <a href="#" class="TileEdit" id="'+GetListItems[i].ID+'" >'+GetListItems[i].Title+'<i class="icon-more" data-toggle="modal" data-target="#Addmodal"></i></a> </div>'+
                   '</div>';
        }
        else if(i==3)
        {
          HTML += '<div class="col-sm-3 col-xs-12 over-right">'+
                    '<div class="img-gallery"><div class="exp-img"><div align="center" class="small-icon"> <img src="'+FavIcon+'"></div></div> <a target="_blank" class="head-h3" href="'+GetListItems[i].LinkURL.Url+'" id="'+GetListItems[i].ID+'" >'+GetListItems[i].Title+'</a> <a class="icon-more TileEdit" style="cursor: pointer;" id="'+GetListItems[i].ID+'" data-toggle="modal" data-target="#Addmodal"></a></div>'+
                    //'<div class="img-gallery"> <img src="images/announce-listimg1.jpg"> <a href="#" class="TileEdit" id="'+GetListItems[i].ID+'" >'+GetListItems[i].Title+'<i class="icon-more" data-toggle="modal" data-target="#Addmodal"></i></a> </div>'+
                  '</div>';
        }
        else
        {
          HTML += '<div class="col-sm-3 col-xs-12">'+
                    '<div class="img-gallery"><div class="exp-img"><div align="center" class="small-icon"> <img src="'+FavIcon+'"></div></div> <a target="_blank" class="head-h3" href="'+GetListItems[i].LinkURL.Url+'" id="'+GetListItems[i].ID+'" >'+GetListItems[i].Title+'</a> <a class="icon-more TileEdit" style="cursor: pointer;" id="'+GetListItems[i].ID+'" data-toggle="modal" data-target="#Addmodal"></a></div>'+
                    //'<div class="img-gallery"> <img src="images/announce-listimg1.jpg"> <a href="#" class="TileEdit" id="'+GetListItems[i].ID+'" >'+GetListItems[i].Title+'<i class="icon-more" data-toggle="modal" data-target="#Addmodal"></i></a> </div>'+
                  '</div>';
        }
      }      
    }
    if(GetListItems.length < 4)
    {
    HTML += '<div class="col-sm-3 col-xs-12 over-right">'+
                '<div class="no-img"> <a href="#" class="AddTile" data-toggle="modal" data-target="#Addmodal"><i class="icon-add"></i>Add Favourites</a></div>'+
              '</div>';
    }
    $('.gallery-section').append(HTML);


    $('.TileEdit').click(function (event){
      $('#myModalLabel').text("EDIT FAVOURITES");
      let TileID = $(this).attr("id");
      $('#btnAddSubmit').hide();
      $('#btnEditSubmit').show();
      $('#btnDelete').show();
      $('#txtTitle').val("");
      $('#txtLinkUrl').val("");
      let GetListItems = readItems(listName, columnArray, 10, "Modified","ID",parseInt(TileID));
      GetListItems.then(result =>{
        $('#txtID').val(result[0].ID);
        $('#txtTitle').val(result[0].Title);
        $('#txtLinkUrl').val(result[0].LinkURL.Url);
      })
    });

    $('.icon-remove').click(function (event){
      $('#txtTitle').val("");
      $('#txtLinkUrl').val("");
      $('#btnAddSubmit').show();
      $('#btnEditSubmit').hide();
    });

    $('.AddTile').click(function (event){
      
      $('#myModalLabel').text("ADD NEW FAVOURITES");
      $('#txtTitle').val("");
      $('#txtLinkUrl').val("");
      $('#btnAddSubmit').show();
      $('#btnEditSubmit').hide();
      $('#Addmodal').show();
      $('#btnDelete').hide();
    });
    
  }

  
  AddNewTile(){
    if ($('.ajs-message').length > 0) {
      $('.ajs-message').remove();
      }
    if(this.Validation())
    {
      $('#btnAddSubmit').hide();
      $('#lblwait').show();
      var listName = "Tiles";
      let itemObj = {
        Title: $('#txtTitle').val(),
        Username: this.context.pageContext.user.displayName,
        LinkURL: {
          "__metadata": {
              "type": "SP.FieldUrlValue"
          },
          Url: $('#txtLinkUrl').val()
        }
      };
      addItems(listName, itemObj).then(result =>{
        location.reload();
      });
    }
  }

  UpdateItem(){
    if ($('.ajs-message').length > 0) {
      $('.ajs-message').remove();
      }
    if(this.Validation())
    {
      
      $('#btnEditSubmit').hide();
      $('#lblwait').show();
      var listName = "Tiles";
      let columnArray: any = ["Title","URL"];
      let itemId =  +$('#txtID').val();
      let itemObj = {
        Title: $('#txtTitle').val(),
        Username: this.context.pageContext.user.displayName,
        LinkURL: {
          "__metadata": {
              "type": "SP.FieldUrlValue"
          },
          Url: $('#txtLinkUrl').val()
        }
      };
      
      updateItem(listName,itemId,itemObj).then(result =>{
        location.reload();
      });
    }
  }

  DeleteItem(){
    var strconfirm="Are you sure you want to delete ?";
    var _this = this;
    alertify.confirm('Confirmation',strconfirm, function (){
        var listName = "Tiles";
        let itemId =  +$('#txtID').val();
        deleteItem(listName,itemId).then(result =>{
          location.reload();
        });
    },function (){
    }).set('closable', false);
  }

  Validation() {
    var regexp = /^(https?|s?ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i
    if (!$('#txtTitle').val()) {
      alertify.set('notifier', 'position', 'top-right');
      alertify.error("Please Enter the Title");
      return false;
    }
    else if(!$('#txtLinkUrl').val()){
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
