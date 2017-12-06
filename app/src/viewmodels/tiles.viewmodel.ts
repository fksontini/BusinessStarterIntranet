// ========================================
// Groups Component View Model
// ========================================
declare function require(name: string);

import { Web } from "sp-pnp-js";

export class TilesViewModel {

    public tiles: KnockoutObservableArray<any> = ko.observableArray([]);
      
    constructor(params: any) {

        let web = new Web(_spPageContextInfo.webAbsoluteUrl); 
        
        let self = this;
       
        web.lists.getByTitle("Länkar").items.top(4).get().then((items) => {
            this.tiles(items);
        });

    }

}







