// ===============================
// Application main entry point
// ===============================
// We must declare this function to get static files like html or CSS via the ts-loader
// More info here https://github.com/TypeStrong/ts-loader
declare var require: {
    <T>(path: string): T;
    (paths: string[], callback: (...modules: any[]) => void): void;
    ensure: (paths: string[], callback: (require: <T>(path: string) => T) => void) => void;
};

require('es6-promise/auto'); // Fix for IE11 (inject the polyfill in the global context)

// View models for components
import { Localization } from "./core/localization";
import { BreadcrumbViewModel } from "./viewmodels/breadcrumb.viewmodel";
import { CarouselViewModel } from "./viewmodels/carousel.viewmodel";
import { GroupsViewModel } from "./viewmodels/groups.viewmodel";
import { TilesViewModel } from "./viewmodels/tiles.viewmodel";
import { ContextualMenuViewModel } from "./viewmodels/contextualmenu.viewmodel";
import { DefaultDisplayTemplateItemViewModel } from "./viewmodels/defaultdisplaytemplateitem.viewmodel";
import { DefaultFilterViewModel } from "./viewmodels/defaultfilter-mui.viewmodel";
import { DocumentDisplayTemplateItemViewModel } from "./viewmodels/documentitem.viewmodel";
import { HeaderLinksViewModel } from "./viewmodels/headerlinks.viewmodel";
import { FooterLinksViewModel } from "./viewmodels/footerlinks.viewmodel";
import { KnockoutComponent } from "./viewmodels/knockoutcomponent";
import { LanguageSwitcherViewModel } from "./viewmodels/languageswitcher.viewmodel";
import { PageDisplayTemplateItemViewModel } from "./viewmodels/pageitem.viewmodel";
import { PageInfoViewModel } from "./viewmodels/pageinfo.viewmodel";
import { PersonaViewModel } from "./viewmodels/persona.viewmodel";
import { SearchBoxViewModel } from "./viewmodels/searchbox.viewmodel";
import { SearchBoxMobileViewModel } from "./viewmodels/searchboxmobile.viewmodel";
import { TopNavViewModel } from "./viewmodels/topnav.viewmodel";
import { TranslationControlViewModel } from "./viewmodels/translationcontrol.viewmodel";
import { ICSCalendarGeneratorViewModel } from "./viewmodels/icscalendargenerator.viewmodel";
import { BotWebChatViewModel } from "./viewmodels/botwebchat.viewmodel";

// Third party libraries
import i18n = require("i18next");
import { Web, ConsoleListener, Logger, LogLevel, setup } from "sp-pnp-js";

// Main style sheet for the application
require("./styles/css/global.scss");
require("./styles/css/layouts.scss");
require("./styles/css/layouts-edit.scss");
require("./styles/css/displaytemplates.scss");

// Images
require("./styles/css/images/spinner.gif");
require("./styles/css/images/default_image.png");
require("./styles/css/images/favicon_intranet.ico");
require("./styles/css/images/flags.png");

// Bootstrap CSS isolation
require("./styles/css/bootstrap/bootstrap-prefix.less");

export class Main {

    // Static methods are mainly used for SharePoint display templates (it is just a public wrapper)
    // We can't use Knockout components here because bindings are not triggered when the display template logic adds the component programmatically
    // We have to apply bindings manually after rendering
    public static initPageDisplayTemplateItemViewModel = (currentItem: any, domElement: string, filterProperty: string, filterValue: string, allLabel: string) => {

        let viewModel = new PageDisplayTemplateItemViewModel(currentItem, filterProperty, filterValue, allLabel);
        ko.applyBindings(viewModel, domElement);
    }

    public static initDocumentDisplayTemplateItemViewModel = (currentItem: any, domElement: string) => {

        let viewModel = new DocumentDisplayTemplateItemViewModel(currentItem);
        ko.applyBindings(viewModel, domElement);
    }

    public static initDefaultDisplayTemplateItemViewModel = (currentItem: any, domElement: string) => {

        let viewModel = new DefaultDisplayTemplateItemViewModel(currentItem);
        ko.applyBindings(viewModel, domElement);
    }

    public static initDefaultMuiFilterViewModel = (domElement: string) => {

        let viewModel = new DefaultFilterViewModel();
        ko.applyBindings(viewModel, domElement);
    }

    public static getResource = (resourceKey: string) => {

        return i18n.t(resourceKey);
    }

    public static jQuery = () => {
        return $;
    }

    /**
     * Register all Knockout components for the entire application
     */
    public registerComponents() {

        // ===============================
        // Register Knockout components   
        // ===============================

        // Component: "MainMenu"
        let mainMenuTemplate = require("./templates/topnav.template.html");
        require("./styles/css/topnav.scss");
        let mainMenuComponent = new KnockoutComponent("component-topnav", TopNavViewModel, mainMenuTemplate);

        // Component: "ContextualMenu"
        let contextualMenuTemplate = require("./templates/contextualmenu.template.html");
        require("./styles/css/contextualmenu.scss");
        let contextualMenuComponent = new KnockoutComponent("component-contextualmenu", ContextualMenuViewModel, contextualMenuTemplate);

        // Component: "Breadcrumb"
        let breadcrumbTemplate = require("./templates/breadcrumb.template.html");
        require("./styles/css/breadcrumb.scss");
        let breadcrumbComponent = new KnockoutComponent("component-breadcrumb", BreadcrumbViewModel, breadcrumbTemplate);

        // Component: "Header" (template only)
        let headerTemplate = require("./templates/header.template.html");
        require("./styles/css/header.scss");
        let headerComponent = new KnockoutComponent("component-header", null, headerTemplate);

        // Component: "Page Info"
        let pageInfoTemplate = require("./templates/pageinfo.template.html");
        require("./styles/css/pageinfo.scss");
        let pageInfoComponent = new KnockoutComponent("component-pageinfo", PageInfoViewModel, pageInfoTemplate);

         // Component: "Persona"
        let personaTemplate = require("./templates/persona.template.html");
        // require("./styles/css/persona.scss");
        let personaComponent = new KnockoutComponent("component-persona", PersonaViewModel, personaTemplate);

        // Component: "Translation Control"
        let translationControlTemplate = require("./templates/translationcontrol.template.html");
        require("./styles/css/translationcontrol.scss");
        let translationcontrolComponent = new KnockoutComponent("component-translationcontrol", TranslationControlViewModel, translationControlTemplate);

        // Component: "Language Switcher"
        let languageSwitcherTemplate = require("./templates/languageswitcher.template.html");
        require("./styles/css/languageswitcher.scss");
        require("./styles/css/flags.scss");
        let languageSwitcherComponent = new KnockoutComponent("component-languageswitcher", LanguageSwitcherViewModel, languageSwitcherTemplate);

        // Component: "Searchbox"
        let searchboxTemplate = require("./templates/searchbox.template.html");
        require("./styles/css/searchbox.scss");
        let searchboxComponent = new KnockoutComponent("component-searchbox", SearchBoxViewModel, searchboxTemplate);

        // Component: "Header Links"
        let headerLinksTemplate = require("./templates/headerlinks.template.html");
        require("./styles/css/headerlinks.scss");
        let headerLinksComponent = new KnockoutComponent("component-headerlinks", HeaderLinksViewModel, headerLinksTemplate);

        // Component: "Footer Links"
        let footerLinksTemplate = require("./templates/footerlinks.template.html");
        require("./styles/css/footerlinks.scss");
        let footerLinksComponent = new KnockoutComponent("component-footerlinks", FooterLinksViewModel, footerLinksTemplate);

        // Component: "Search Box (mobile)"
        let searchboxMobileTemplate = require("./templates/searchboxmobile.template.html");
        require("./styles/css/searchboxmobile.scss");
        let searchboxMobileComponent = new KnockoutComponent("component-searchboxmobile", SearchBoxMobileViewModel, searchboxMobileTemplate);

        // Component: "Carousel"
        let carouselTemplate = require("./templates/carousel.template.html");
        require("./styles/css/carousel.scss");
        let carouselComponent = new KnockoutComponent("component-carousel", CarouselViewModel, carouselTemplate);

        // Component: "Groups"
        let groupsTemplate = require("./templates/groups.template.html");
        require("./styles/css/groups.scss");
        let groupsComponent = new KnockoutComponent("component-groups", GroupsViewModel, groupsTemplate);

         // Component: "Tiles"
        let tilesTemplate = require("./templates/tiles.template.html");
        require("./styles/css/tiles.scss");
        let tilesComponent = new KnockoutComponent("component-tiles", TilesViewModel, tilesTemplate);

        // Component: "ICS Generator"
        let calendarGeneratorTemplate = require("./templates/icscalendargenerator.html");
        require("./styles/css/icscalendargenerator.scss");
        let calendarGeneratorComponent = new KnockoutComponent("component-icsgenerator", ICSCalendarGeneratorViewModel, calendarGeneratorTemplate);

        // Component: "Bot Web chat"
        let botWebChatTemplate = require("./templates/botwebchat.html");
        require("./styles/css/botwebchat.scss");
        let botWebChatComponent = new KnockoutComponent("component-botwebchat", BotWebChatViewModel, botWebChatTemplate);        
    }

    public init() {

        this.registerComponents();

        // Init the loggger
        let consoleListener = new ConsoleListener();
        Logger.subscribe(consoleListener);
        Logger.activeLogLevel = LogLevel.Error;

        // Needed for SharePoint 2013 On-Premise othjerwise it will use Atom XML
        setup({
            headers: {
                Accept: "application/json; odata=verbose",
            },
        });

        // Be careful, we need to apply bindings after the document is ready
        $(document).ready(() => {

            let localization = new Localization();
                
            localization.initLanguageEnv().then(() => {  

                let web = new Web(_spPageContextInfo.webAbsoluteUrl);         

                // Apply the Knockout JS magic!
                ko.applyBindings();

                // Add Bootstrap responsive behavior for news images
                $("#page-image img").addClass("img-responsive");

                web.lists.getByTitle("Sidor").items.getById(_spPageContextInfo.pageItemId).select("HideSideBar").get().then(item => {
                    
                    if (item.HideSideBar) {

                        $("#sidebar").hide();
                        $("#content").removeClass("col-md-push-3 col-md-9");
                        $("#content").addClass("col-md-12");

                        $("#breadcrumb-nav").hide();
                        $(".page-layout #title").hide();
                    }
                });            
            });            
        });
    }
}

// Start the engine
let main = new Main();
main.init();
