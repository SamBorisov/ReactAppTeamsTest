import React from 'react'
import * as microsoftTeams from "@microsoft/teams-js"
import "../lib/auth"
import { MsalAuthService } from "../lib/msal-auth"
import $ from 'jquery';
// SSO --------------------------------------------------

export default function SSO() {


    // () => {
    // microsoftTeams.initialize();
    // let authTokenRequestOptions = {
    //     successCallback: (result) => { this.ssoLoginSuccess(result) },
    //     failureCallback: (error) => { this.ssoLoginFailure(error) }
    // };

    // microsoftTeams.authentication.getAuthToken(authTokenRequestOptions);
    // }
    const authService = new MsalAuthService("c721dec4-628c-42b4-aecf-5376bd70435a", "api://teamstabsso2022122.azurewebsites.net/c721dec4-628c-42b4-aecf-5376bd70435a");

    $(document).ready(function() {

        authService
            .isCallback()
            .then((isCallback) => {
                if (!isCallback) {
                    authService
                        .getUser()
                        .then((user) => {
                            // Signed in the user automatically; we're ready to go
                            setUserSignedIn(true);
                            getMyProfile(user.localAccountId);
                        })
                        .catch(() => {
                            setUserSignedIn(false);
                            // Failed to sign in the user automatically; show login screen
                            console.log("Failed")
                        });
                }
            })
            .catch((error) => {
                // Encountered error during redirect login flow; show error screen
                console.log(error);
            });
    });

    function login() {
        authService
            .login()
            .then((user) => {
                if (user) {
                    setUserSignedIn(true);
                    getMyProfile(user.localAccountId);
                } else {
                    setUserSignedIn(false);
                }
            })
            .catch((err) => {
                setUserSignedIn(false);
                console.error(err);
            });
    };

    function logout() {
        authService.logout();
    }

    function getMyProfile(userId) {
        setUserSignedIn(true);
        authService.getUserInfo(userId);
       // debugger;
        authService.getToken()
            .then(data => {
                getSPOInformation(data); // Get SharePoint information
            });
       
    }

    function setUserSignedIn(isUserSignedIn) {
        document.getElementById("browser-login").hidden = isUserSignedIn;
    }

    function getSPOInformation(token){

        /// check if token undefinend
        if (token == undefined) {
            return;
        }

            
        //const functionUrl = `https://teamsalex002.loophole.site/api/ConsumeSPO`;
        const functionUrl = `https://piasysyoteamsssobackend20221219233441.azurewebsites.net/api/ConsumeSPO?code=1SAHSE1Hi7vA577n5Y2su6vQGewH8sLaP-sx6fGv6meuAzFuGaMYLw==`;
        fetch(functionUrl, {
          headers: {
                Accept: 'application/json',
                Authorization: 'Bearer ' + token
          }
        })
           .then(resp => resp.json())
           .then(json => {
                document.getElementById("divUPN").textContent = json.userPrincipalName;
                document.getElementById("divID").textContent = json.id;
                document.getElementById("divName").textContent = json.name;
                console.log(json);
           }
           )
        
    }

    



    return (
        <div className='container'>

                <h1>Tab SSO Authentication</h1>
                <div id="browser-signin-text" style="display:none;">Try signing in from browser - <a href="/Home/">Click here</a></div>
                <div id="divError" style="display: none"></div>
                <button onclick="requestConsent()" id="consent" style="display:none;">Authenticate</button>
                <div id="divGraphProfile" style="display: none"></div>
                <div id="browser-signin-container" style="display:none;">
                    <div id="browser-login">
                        <h2>Please click on Login button to see your profile details!</h2>
                        <button class="browser-button" onclick="login()">Login</button>
                    </div>  
                </div>
                <div>id: <span id="divID"></span></div>
                <div>UPN: <span id="divUPN"></span></div>
                <div>Name: <span id="divName"></span></div>                       
                   
            
        </div>
        )

}