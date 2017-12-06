// ========================================
// Bot web chat View Model
// ========================================
import { PersonaComponent } from "../components/persona"; 
import * as React from "react";
import * as ReactDOM from "react-dom";

export class PersonaViewModel {
    private userEmail: any;
    private element: any;
    constructor(params: any) {
        debugger;
         this.userEmail = params.userEmail;
         this.element = params.element;
         
        // We encapsulate the React component in a Knockout component to be able to control the DOM anchor point.
        // If you call the render() method directly in the main.ts, it means the element with id 'bot-webchat' has to be present in the master page initially (error otherwise).
        ReactDOM.render(<PersonaComponent email={this.userEmail} />, document.getElementById(this.element)); 
    }
}

                
                

