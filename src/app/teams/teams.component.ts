import { Component, OnInit } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { HttpClient } from '@angular/common/http';
import { InteractionRequiredAuthError, AuthError } from 'msal';

const GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0/me';

@Component({
  selector: 'app-teams',
  templateUrl: './teams.component.html',
  styleUrls: ['./teams.component.css']
})
export class TeamsComponent implements OnInit {
  profile;
  profile2json;
  accessToken: string;

  constructor(private authService: MsalService, private http: HttpClient) { }

  ngOnInit() {
    this.getProfile();
  }

  getProfile() {
    this.http.get(GRAPH_ENDPOINT)
    .subscribe({
      next: (profile) => {
        this.profile = profile;
        this.profile2json = JSON.stringify(profile);
        this.accessToken = "";
      },
      error: (err: AuthError) => {
        // If there is an interaction required error,
        // call one of the interactive methods and then make the request again.
        if (InteractionRequiredAuthError.isInteractionRequiredError(err.errorCode)) {

          this.authService.acquireTokenPopup({
            scopes: this.authService.getScopesForEndpoint(GRAPH_ENDPOINT)
          })
          .then(function (tokenResponse) {

            console.log(tokenResponse.accessToken);

            this.http.get(GRAPH_ENDPOINT)
              .toPromise()
              .then(profile => {
                this.profile = profile;
              });
          });
        }
      }
    });
  }

  showAccessToken() {
    var caller = this;

    if (this.accessToken == "")
    {
      const requestObj = {
          scopes: ["user.read"]
      };

      this.authService.acquireTokenSilent(requestObj).then(function (tokenResponse) {
          // Callback code here
          caller.accessToken = tokenResponse.accessToken;
          }).catch(function (error) {
              console.log(error);
          });
    }
    else
    {
      this.accessToken = "";
    }
  }
}