import { Component } from '@angular/core';

import { NavController } from 'ionic-angular';

import Microsoft from '../../../plugins/cordova-plugin-ms-adal/cordova-plugin-ms-adal.d.ts'

@Component({
  selector: 'page-about',
  templateUrl: 'about.html'
})
export class AboutPage {

  constructor(public navCtrl: NavController) {

    console.log('constructor')

      var resourceUrl = 'https://outlook.office365.com';
      var officeEndpointUrl = 'https://outlook.office365.com/ews/odata';
      var appId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx';
      var authUrl = 'https://login.windows.net/common/';
      var redirectUrl = 'http://localhost:4400/services/office365/redirectTarget.html';


      console.log('constructor: 1')
      var AuthenticationContext = Microsoft.ADAL.AuthenticationContext;

      console.log('constructor: 2')
      var outlookClient = new Microsoft.OutlookServices.Client(officeEndpointUrl,
        new AuthenticationContext(authUrl), resourceUrl, appId, redirectUrl);

        console.log('sup mother');
        console.log(outlookClient);

      outlookClient.me.folders.getFolder('Inbox').messages.getMessages().fetchAll().then(function (result) {
        result.forEach(function (msg) {
            console.log('Message "' + msg.Subject + '" received at "' + msg.DateTimeReceived.toString() + '"');
        });
      }, function(error) {
        console.error(error);
      });

  }

  ngOnInit() {
  }

}
