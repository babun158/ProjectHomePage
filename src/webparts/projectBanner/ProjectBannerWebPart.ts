import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ProjectBannerWebPart.module.scss';
import * as strings from 'ProjectBannerWebPartStrings';
import  'jquery';
import pnp from 'sp-pnp-js'
import { readItems, checkUserinGroup } from '../../commonJS';

declare var $;
export interface IProjectBannerWebPartProps {
  description: string;
}

export default class ProjectBannerWebPart extends BaseClientSideWebPart<IProjectBannerWebPartProps> {

  userflag: boolean = false;
  public render(): void {
    var _this = this;
    
    //Checking user details in group
    
    checkUserinGroup("Banners", this.context.pageContext.user.email, function (result) {
      if (result == 1) {
        _this.userflag = true;
      }
      _this.viewlistitemdesign();
    })
   
  }

  public viewlistitemdesign(){
    var siteURL = this.context.pageContext.web.absoluteUrl;
    this.domElement.innerHTML = 
    '<section class="banner-section">' +
    "<div class='ban-section'>"+
    '<h3 class="tt-head">'+ this.context.pageContext.web.title +'</h3>' +
      '<div id="carousel-banner" class="carousel carousel-fade" data-ride="carousel">' +
        '<ol class="carousel-indicators banner-carousel">' +
        '</ol>' +
      '</div>' +
      '<div id="addEvents" class="event-add" style="Display:none">' +
        //'<h3 class="banner-itemview" style="cursor:pointer">UPDATE COVERAGE EVENTS <a href="' + siteURL + '/Pages/ListView.aspx?CName=Banners"></a><i href="' + siteURL + '/Pages/AddListItem.aspx?CName=Banners" class="icon-add"></i></h3>' +
        '<h3 class="banner-itemview" style="cursor:pointer">VIEW COVERAGE EVENTS <a href="' + siteURL + '/Pages/AddListItem.aspx?CName=Banners"><i class="icon-add"></i></a> </h3>' +
      '</div>' +
    "</div>"+
    '<section>';  ;   
    this.BannerPage(this.userflag);
    let viewevent = document.getElementsByClassName('banner-itemview');
    for (let i = 0; i < viewevent.length; i++) {
      viewevent[i].addEventListener("click", (e: Event) => this.viewpageRedirect(siteURL));
    }
    $('#carousel-banner').carousel({ interval: 3000 });
  }
  viewpageRedirect(siteURL){
    window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=Banners";
  }
  async BannerPage(userflag) {
    var renderhtml = '<div class="carousel-inner" role="listbox">';
    var renderliitems = "";
    var count;
    var activeflag;
    let bannerItems;
    var siteURL = this.context.pageContext.site.absoluteUrl;
    bannerItems = await readItems("Banners", ["Title", "Modified", "LinkURL", "Display", "BannerContent", "Image"], 3, "Modified", "Display", 1);
    if(bannerItems.length>0)
    {
      for (let i = 0; i < bannerItems.length; i++) { 
        if (i == 0) {
          activeflag = "active";
        } else {
          activeflag = "";
        }
        renderliitems += '<li data-slide-to="' + i + '" data-target="#carousel-banner" class="' + activeflag + '">' + '</li>';
        renderhtml += '<div class="item ' + activeflag + '">';
        renderhtml += '<img src="' + bannerItems[i].Image.Url + '" style="max-height: 319px;"alt="Slide" title="Slide" />';
        renderhtml += '<div class="carousel-caption">';
        var DottedTitle=bannerItems[i].Title;
        if(DottedTitle.length>65)
        {
          DottedTitle=DottedTitle.substring(0,65)+"...";
        }
        if (bannerItems[i].LinkURL !== null) {
          renderhtml += '<div align="center">' + '<a href="' + bannerItems[i].LinkURL.Url + '" class="wow fadeInRight" style="visibility: visible; animation-name: fadeInRight;">lEARN mORE</a>' + '</div>';
        }
        renderhtml += '</div>';
        renderhtml += '</div>';
      }
    }
      else if(bannerItems.length==0){
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
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
