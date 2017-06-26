import com.google.api.client.extensions.java6.auth.oauth2.VerificationCodeReceiver;
import com.google.api.client.util.Throwables;
import org.mortbay.jetty.Connector;
import org.mortbay.jetty.Request;
import org.mortbay.jetty.Server;
import org.mortbay.jetty.handler.AbstractHandler;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.net.Socket;
import java.net.SocketAddress;
import java.util.concurrent.locks.Condition;
import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReentrantLock;

/**
 * Custom implementation of "LocalServerReceiver" to override the callback port
 */
public class CustomServerReceiver implements VerificationCodeReceiver {

    private Server server;
    String code;
    String error;
    final Lock lock;
    final Condition gotAuthorizationResponse;
    private final int port;
    private final String host;

    public CustomServerReceiver() {
        this("localhost", 4568);
    }

    CustomServerReceiver(String host, int port) {
        this.lock = new ReentrantLock();
        this.gotAuthorizationResponse = this.lock.newCondition();
        this.host = host;
        this.port = port;
    }

    public String getRedirectUri() throws IOException {

        this.server = new Server(this.port);
        Connector[] e = this.server.getConnectors();
        int len$ = e.length;

        for(int i$ = 0; i$ < len$; ++i$) {
            Connector c = e[i$];
            c.setHost(this.host);
        }

        this.server.addHandler(new CustomServerReceiver.CallbackHandler());

        try {
            this.server.start();
        } catch (Exception var5) {
            Throwables.propagateIfPossible(var5);
            throw new IOException(var5);
        }

        return "http://" + this.host + ":" + this.port + "/callback";
    }

    public String waitForCode() throws IOException {
        this.lock.lock();

        String var1;
        try {
            while(this.code == null && this.error == null) {
                this.gotAuthorizationResponse.awaitUninterruptibly();
            }

            if(this.error != null) {
                throw new IOException("User authorization failed (" + this.error + ")");
            }

            var1 = this.code;
        } finally {
            this.lock.unlock();
        }

        return var1;
    }

    public void stop() throws IOException {
        if(this.server != null) {
            try {
                this.server.stop();
            } catch (Exception var2) {
                Throwables.propagateIfPossible(var2);
                throw new IOException(var2);
            }

            this.server = null;
        }

    }

    public String getHost() {
        return this.host;
    }

    public int getPort() {
        return this.port;
    }

    class CallbackHandler extends AbstractHandler {
        CallbackHandler() {
        }

        public void handle(String target, HttpServletRequest request, HttpServletResponse response, int dispatch) throws IOException {
            if("/callback".equals(target)) {
                this.writeLandingHtml(response);
                response.flushBuffer();
                ((Request)request).setHandled(true);
                CustomServerReceiver.this.lock.lock();

                try {
                    CustomServerReceiver.this.error = request.getParameter("error");
                    CustomServerReceiver.this.code = request.getParameter("code");
                    CustomServerReceiver.this.gotAuthorizationResponse.signal();
                } finally {
                    CustomServerReceiver.this.lock.unlock();
                }

            }
        }

        private void writeLandingHtml(HttpServletResponse response) throws IOException {
            response.setStatus(200);
            response.setContentType("text/html");
            PrintWriter doc = response.getWriter();
            doc.println("<html>");
            doc.println("<head><title>OAuth 2.0 Authentication Token Received</title></head>");
            doc.println("<body>");
            doc.println("Received verification code. You may now close this window...");
            doc.println("</body>");
            doc.println("</HTML>");
            doc.flush();
        }
    }

    public static final class Builder {
        private String host = "localhost";
        private int port = -1;

        public Builder() {
        }

        public CustomServerReceiver build() {
            return new CustomServerReceiver(this.host, this.port);
        }

        public String getHost() {
            return this.host;
        }

        public CustomServerReceiver.Builder setHost(String host) {
            this.host = host;
            return this;
        }

        public int getPort() {
            return this.port;
        }

        public CustomServerReceiver.Builder setPort(int port) {
            this.port = port;
            return this;
        }
    }
}