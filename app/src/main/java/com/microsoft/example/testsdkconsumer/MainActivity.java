package com.microsoft.example.testsdkconsumer;

import android.content.Intent;
import android.os.Bundle;
import android.support.v7.app.AppCompatActivity;
import android.view.View;
import android.widget.TextView;

import com.microsoft.aad.adal.AuthenticationCallback;
import com.microsoft.aad.adal.AuthenticationContext;
import com.microsoft.aad.adal.AuthenticationResult;
import com.microsoft.aad.adal.PromptBehavior;
import com.microsoft.example.testsdkconsumer.app.App;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.extensions.IGraphServiceClient;
import com.microsoft.graph.extensions.IMessageCollectionPage;
import com.microsoft.graph.extensions.Message;

import java.io.PrintWriter;
import java.io.StringWriter;

import butterknife.BindView;
import butterknife.ButterKnife;
import butterknife.OnClick;
import timber.log.Timber;

import static com.microsoft.example.testsdkconsumer.R.id.btn_connect;
import static com.microsoft.example.testsdkconsumer.R.id.btn_makereq;
import static com.microsoft.example.testsdkconsumer.R.id.txt_dat;
import static com.microsoft.example.testsdkconsumer.R.layout.activity_main;
import static com.microsoft.example.testsdkconsumer.app.App.CLIENT_ID;
import static com.microsoft.example.testsdkconsumer.app.App.GRAPH_RESOURCE;
import static com.microsoft.example.testsdkconsumer.app.App.REDIRECT_URI;

public class MainActivity extends AppCompatActivity {

    // You could dependency-inject these fields into place for cleaner code
    protected App mApplication;
    protected AuthenticationContext mAuthenticationContext;
    protected IGraphServiceClient mGraphServiceClient;

    @BindView(txt_dat)
    protected TextView mDat;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(activity_main);

        // grab our application
        mApplication = (App) getApplication();

        // get an instance of our authcontext for sign-in
        mAuthenticationContext = mApplication.getAuthenticationContext();

        // grab the gsc for making calls to the graph
        mGraphServiceClient = mApplication.getGraphServiceClient();

        // set up butterknife
        ButterKnife.bind(this);
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        mAuthenticationContext.onActivityResult(requestCode, resultCode, data);
    }

    /**
     * Called when the connect button is clicked - this impl performs auth
     *
     * @param v the View which was clicked
     */
    @OnClick(btn_connect)
    public void onConnectClicked(View v) {
        mAuthenticationContext
                .acquireToken(
                        this,
                        GRAPH_RESOURCE,
                        CLIENT_ID,
                        REDIRECT_URI,
                        PromptBehavior.Auto,
                        new AuthenticationCallback<AuthenticationResult>() {
                            @Override
                            public void onSuccess(AuthenticationResult result) {
                                Timber.d("Success");
                                mDat.setText("Authenticated");
                                mApplication
                                        .setGraphUserId(
                                                result.getUserInfo().getUserId()
                                        );
                            }

                            @Override
                            public void onError(Exception exc) {
                                Timber.e(exc, "Error");
                                displayThrowable(exc);
                            }
                        }
                );
    }

    /**
     * Called when Make Request is clicked - this impl loads some emails
     *
     * @param v the View which was clicked
     */
    @OnClick(btn_makereq)
    public void onMakeReqClicked(View v) {
        // Make a test call to the service - in this case, I'm going to fetch email
        mGraphServiceClient
                .getMe()
                .getMessages()
                .buildRequest()
                .get(new ICallback<IMessageCollectionPage>() {
                    @Override
                    public void success(IMessageCollectionPage iMessageCollectionPage) {
                        String msg = "Loaded emails:\n";
                        for (Message m : iMessageCollectionPage.getCurrentPage()) {
                            msg += m.subject + "\n";
                        }
                        mDat.setText(msg);
                    }

                    @Override
                    public void failure(ClientException ex) {
                        Timber.e(ex, "Failed to load messages");
                        // Set the TextView to display the Exception
                        displayThrowable(ex);
                    }
                });
    }

    private void displayThrowable(Throwable throwable) {
        StringWriter sw = new StringWriter();
        PrintWriter pw = new PrintWriter(sw);
        throwable.printStackTrace(pw);
        String trace = sw.toString();
        mDat.setText(trace);
    }
}
