package com.microsoft.example.testsdkconsumer.app;

import android.app.Application;

import com.microsoft.aad.adal.AuthenticationContext;
import com.microsoft.aad.adal.AuthenticationException;
import com.microsoft.aad.adal.AuthenticationResult;
import com.microsoft.graph.authentication.IAuthenticationProvider;
import com.microsoft.graph.extensions.GraphServiceClient;
import com.microsoft.graph.extensions.IGraphServiceClient;
import com.microsoft.graph.http.IHttpRequest;

import timber.log.Timber;

import static com.microsoft.graph.core.DefaultClientConfig.createWithAuthenticationProvider;

public class App extends Application {

    // Keep these secret
    public static final String CLIENT_ID = "YOUR CLIENT ID";
    public static final String REDIRECT_URI = "YOUR REDIRECT URI";

    // The resource id is a uri for the service we wish to access
    public static final String GRAPH_RESOURCE = "https://graph.microsoft.com";

    public static final String AUTHORITY_URL = "https://login.microsoftonline.com/common";

    // Our Graph AuthenticationContext
    private AuthenticationContext mGraphAuthenticationContext;

    // Our SDK instance
    private IGraphServiceClient mGSC;

    // Store this value somewhere like a db
    // or sharedprefs unless you want it to go away at app restarts
    private String mGraphUserId;

    @Override
    public void onCreate() {
        super.onCreate();
        Timber.plant(new Timber.DebugTree());
        mGraphAuthenticationContext = new AuthenticationContext(
                this, // our application's context
                AUTHORITY_URL, // our authority url
                true // validate the authority
        );
    }

    private IAuthenticationProvider getGraphIAuthenticationProvider() {
        return new IAuthenticationProvider() {
            @Override
            public void authenticateRequest(IHttpRequest request) {
                // Should only get called on a bg-thread
                try {
                    AuthenticationResult authenticationResult = mGraphAuthenticationContext
                            .acquireTokenSilentSync(
                                    GRAPH_RESOURCE,
                                    CLIENT_ID,
                                    getGraphUserId()
                            );
                    Timber.d("Adding auth header");
                    String accessToken = authenticationResult.getAccessToken();
                    request.addHeader("Authorization", "Bearer " + accessToken);
                } catch (AuthenticationException | InterruptedException e) {
                    e.printStackTrace();
                    // calls will fail with 401
                }
            }
        };
    }

    /**
     * Stashes the userid after auth - ideally this hits some DAO for safekeeping
     *
     * @param userId the authenticated user's userid
     */
    public void setGraphUserId(String userId) {
        mGraphUserId = userId;
    }

    /**
     * Gets the stashed user id from the last auth - ideally this hits some DAO
     *
     * @return the stashed user id
     */
    public String getGraphUserId() {
        return mGraphUserId;
    }

    /**
     * Gets the Graph AuthenticationContext
     *
     * @return the Application's AuthenticationContext
     */
    public AuthenticationContext getAuthenticationContext() {
        return mGraphAuthenticationContext;
    }

    /**
     * Gets the GraphServiceClient - the SDK object with which we'll hit the service
     *
     * @return a GraphServiceClient instance
     */
    public IGraphServiceClient getGraphServiceClient() {
        if (null == mGSC) {
            mGSC = new GraphServiceClient
                    .Builder()
                    .fromConfig(
                            createWithAuthenticationProvider(
                                    getGraphIAuthenticationProvider()
                            )
                    ).buildClient();
        }
        return mGSC;
    }
}
