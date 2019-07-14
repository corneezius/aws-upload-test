import { Component } from '@angular/core';


const template = require('./app.component.html');

@Component({
    selector: 'app-home',
    template
})
export default class AppComponent {
    welcomeMessage = 'Welcome';

    async inputChanged(event) {
        return Word.run(async context => {
            /**
             * Insert your Word code here
             */
            const thisDocument = context.document;
            const stateControls = thisDocument.contentControls.getByTag(event.target.name);
            stateControls.load("items");
            context.sync().then( () => {
                stateControls.items.forEach(thisControl => {
                    thisControl.insertText(event.target.value, Word.InsertLocation.replace);
                });
                return context.sync();
            });

            await context.sync();
        });
    }

    CreateNewDocument(event) {
        Word.run(async context => {
          try {
              var newDoc = context.application.createDocument('');
              newDoc.open();
              context.sync();
          } catch (error) {
              console.log(error);
          }
          event.completed();
          context.sync();
        });
    }
}