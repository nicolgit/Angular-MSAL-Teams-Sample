import { Component, OnInit } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { HttpClient } from '@angular/common/http';
import { InteractionRequiredAuthError, AuthError } from 'msal';
import { CreateMeetingResponse } from '../model/CreateMeetingResponse';

const GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0/me';
const GRAPH_TEAMSENDPOINT = 'https://graph.microsoft.com/v1.0/me/onlineMeetings';

@Component({
  selector: 'app-teams',
  templateUrl: './teams.component.html',
  styleUrls: ['./teams.component.css']
})
export class TeamsComponent implements OnInit {
  profile;
  profile2json: string;

  result;
  result2json: string;

  meetingURL: string;

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

  createTeamsMeeting() {
    var caller = this;

    this.http.post(GRAPH_TEAMSENDPOINT, { })
    .subscribe({
      next: (result: CreateMeetingResponse) => {
        this.result = result;
        this.result2json = JSON.stringify(result);
        this.meetingURL = result.joinUrl;
      },
      error: (err: AuthError) => {
        // If there is an interaction required error,
        // call one of the interactive methods and then make the request again.
        if (InteractionRequiredAuthError.isInteractionRequiredError(err.errorCode)) {

          this.authService.acquireTokenPopup({
            scopes: this.authService.getScopesForEndpoint(GRAPH_TEAMSENDPOINT)
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
}