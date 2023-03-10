const {
    Client
} = require("@microsoft/microsoft-graph-client");
const {
    TokenCredentialAuthenticationProvider
} = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
const {
    AuthorizationCodeCredential
} = require("@azure/identity");


const credential = new AuthorizationCodeCredential(
    "69c587b3-8640-4e28-aec8-4efbd62532d8",
    "9f6ad88b-0b32-4c32-aa35-ba8dcd288995",
    "OI28Q~RAu4bQAAoVL_nZmNjtZxCiTzAw28pHUaMi",
    "http://localhost:3007"
);
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ["User.Read"]
});



const client = Client.initWithMiddleware({
    debugLogging: true,
    authProvider
    // Use the authProvider object to create the class.
});