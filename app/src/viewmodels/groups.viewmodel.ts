// ========================================
// Groups Component View Model
// ========================================
declare function require(name: string);

import { Web } from "sp-pnp-js";
import "trunk8";
import * as i18n from "i18next";
let Flickity = require('flickity');
require('flickity-imagesloaded');

export class GroupsViewModel {

    public groups: KnockoutObservableArray<any> = ko.observableArray([]);
    // public siteLogoUrl: KnockoutObservable<string> = ko.observable("");
    // private readMoreLabel: KnockoutObservable<string> = ko.observable("");
    
    constructor(params: any) {
   
        let trunk8OptionsNavLabel: Trunk8Options = {
            lines: 2,
            tooltip: false,
        };

        let trunk8OptionsSlideTitle: Trunk8Options = {
            lines: 4,
            tooltip: false,
        };

        let web = new Web("https://dizparc.sharepoint.com/sites/intranet");
        
        let self = this;
        // Get the current page language
        web.lists.getByTitle("Groups").items.get().then((items) => {

            this.groups(items);

            this.fixPaging();

            $("#filter").keyup(function () {
                var filter = $(this).val(), count = 0;

                // Loop through the comment list
                $("#test li").each(function () {
                    if ($(this).text().search(new RegExp(filter, "i")) < 0) {
                        $(this).fadeOut();

                    } else {
                        $(this).show();
                        count++;
                    }
                });

                // Update the count
                var numberItems = count;
                self.fixPaging();

            });

        });

    }

    public fixPaging = () => {
        var totalContent = $('#test li:visible').length;
        var onePageContent = 12;

        //Page number and Math.round for balancing page
        $("#test li:gt(" + (onePageContent - 1) + ")").hide();
        var totalPage = Math.round(totalContent / onePageContent);

        $(".page").empty();
        for (let i = 1; i <= totalPage; i++) {
            $(".page").append("<a href='javascript:void(0)'>" + i + "</a>");
        }

        // first page added active class
        $(".page a:first").addClass("active");

        // click function
        $(".page a").on("click", (event) => {
            let element = $(event.target);
            let index = $(element).index() + 1;
            let gt = onePageContent * index;

            $(".page a").removeClass("active");
            $(element).addClass("active");
            $("#test li").hide();

            for (let i = gt - onePageContent; i < gt; i++) {
                $("#test li:eq(" + i + ")").show();
            }

        });

    }


}







