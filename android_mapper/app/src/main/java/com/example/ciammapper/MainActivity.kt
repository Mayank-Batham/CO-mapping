package com.example.ciammapper

import android.annotation.SuppressLint
import android.app.DownloadManager
import android.content.Context
import android.content.Intent
import android.net.ConnectivityManager
import android.net.NetworkCapabilities
import android.net.Uri
import android.net.wifi.WifiManager
import android.os.Build
import android.os.Bundle
import android.os.Environment
import android.view.View
import android.webkit.*
import android.widget.Button
import android.widget.EditText
import android.widget.LinearLayout
import android.widget.TextView
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import kotlinx.coroutines.*
import java.net.InetSocketAddress
import java.net.Socket
import java.util.Locale

class MainActivity : AppCompatActivity() {

    private lateinit var webView: WebView
    private lateinit var loadingOverlay: LinearLayout
    private lateinit var statusText: TextView
    private lateinit var errorOverlay: LinearLayout
    private lateinit var btnRetry: Button
    private lateinit var btnResetServer: Button

    // Dynamic Server Setup Views
    private lateinit var urlConfigOverlay: LinearLayout
    private lateinit var etServerUrl: EditText
    private lateinit var btnConnect: Button
    private lateinit var btnScanLocal: Button

    private var fileUploadCallback: ValueCallback<Array<Uri>>? = null

    private val PREFS_NAME = "CIAMapperPrefs"
    private val KEY_SERVER_URL = "server_url"
    private val DEFAULT_CLOUD_URL = "https://your-app-name.streamlit.app"

    // Register file chooser launcher for WebView uploads
    private val fileChooserLauncher = registerForActivityResult(
        ActivityResultContracts.StartActivityForResult()
    ) { result ->
        if (result.resultCode == RESULT_OK) {
            val data: Intent? = result.data
            val results = if (data != null) {
                val dataString = data.dataString
                val clipData = data.clipData
                if (clipData != null) {
                    val uris = Array(clipData.itemCount) { i -> clipData.getItemAt(i).uri }
                    uris
                } else if (dataString != null) {
                    arrayOf(Uri.parse(dataString))
                } else {
                    null
                }
            } else {
                null
            }
            fileUploadCallback?.onReceiveValue(results)
        } else {
            fileUploadCallback?.onReceiveValue(null)
        }
        fileUploadCallback = null
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        webView = findViewById(R.id.webView)
        loadingOverlay = findViewById(R.id.loadingOverlay)
        statusText = findViewById(R.id.statusText)
        errorOverlay = findViewById(R.id.errorOverlay)
        btnRetry = findViewById(R.id.btnRetry)
        btnResetServer = findViewById(R.id.btnResetServer)

        urlConfigOverlay = findViewById(R.id.urlConfigOverlay)
        etServerUrl = findViewById(R.id.etServerUrl)
        btnConnect = findViewById(R.id.btnConnect)
        btnScanLocal = findViewById(R.id.btnScanLocal)

        setupWebView()

        btnRetry.setOnClickListener {
            checkAndLoadUrl()
        }

        btnResetServer.setOnClickListener {
            showUrlConfigScreen()
        }

        btnConnect.setOnClickListener {
            handleConnect()
        }

        btnScanLocal.setOnClickListener {
            startDiscovery()
        }

        checkAndLoadUrl()
    }

    @SuppressLint("SetJavaScriptEnabled")
    private fun setupWebView() {
        val settings = webView.settings
        settings.javaScriptEnabled = true
        settings.domStorageEnabled = true
        settings.allowFileAccess = true
        settings.allowContentAccess = true
        settings.useWideViewPort = true
        settings.loadWithOverviewMode = true
        
        // Handle client behavior
        webView.webViewClient = object : WebViewClient() {
            override fun onPageFinished(view: WebView?, url: String?) {
                super.onPageFinished(view, url)
                // Dismiss splash screen and show WebView
                loadingOverlay.visibility = View.GONE
                webView.visibility = View.VISIBLE
            }
            
            override fun onReceivedError(
                view: WebView?,
                request: WebResourceRequest?,
                error: WebResourceError?
            ) {
                super.onReceivedError(view, request, error)
                // Only show connection error screen if the main webpage loading fails
                if (request?.isForMainFrame == true) {
                    showErrorState()
                }
            }
        }

        // Custom WebChromeClient to support file uploads in WebView
        webView.webChromeClient = object : WebChromeClient() {
            override fun onShowFileChooser(
                webView: WebView?,
                filePathCallback: ValueCallback<Array<Uri>>?,
                fileChooserParams: FileChooserParams?
            ): Boolean {
                fileUploadCallback?.onReceiveValue(null)
                fileUploadCallback = filePathCallback

                val intent = Intent(Intent.ACTION_GET_CONTENT).apply {
                    addCategory(Intent.CATEGORY_OPENABLE)
                    type = "*/*"
                    // Restrict user selection to office and excel mime types
                    val mimeTypes = arrayOf(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // .xlsx
                        "application/vnd.ms-excel", // .xls
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document", // .docx
                        "application/msword" // .doc
                    )
                    putExtra(Intent.EXTRA_MIME_TYPES, mimeTypes)
                    putExtra(Intent.EXTRA_ALLOW_MULTIPLE, true)
                }

                try {
                    fileChooserLauncher.launch(Intent.createChooser(intent, "Select Excel / QP files"))
                } catch (e: Exception) {
                    fileUploadCallback?.onReceiveValue(null)
                    fileUploadCallback = null
                    Toast.makeText(this@MainActivity, "Cannot open file picker", Toast.LENGTH_SHORT).show()
                    return false
                }
                return true
            }
        }

        // Setup File Downloads Support in WebView
        webView.setDownloadListener { url, userAgent, contentDisposition, mimetype, contentLength ->
            try {
                val request = DownloadManager.Request(Uri.parse(url)).apply {
                    setMimeType(mimetype)
                    addRequestHeader("User-Agent", userAgent)
                    setDescription("Downloading Consolidated Excel...")
                    setTitle(URLUtil.guessFileName(url, contentDisposition, mimetype))
                    setNotificationVisibility(DownloadManager.Request.VISIBILITY_VISIBLE_NOTIFY_COMPLETED)
                    setDestinationInExternalPublicDir(
                        Environment.DIRECTORY_DOWNLOADS,
                        URLUtil.guessFileName(url, contentDisposition, mimetype)
                    )
                }

                val dm = getSystemService(DOWNLOAD_SERVICE) as DownloadManager
                dm.enqueue(request)
                Toast.makeText(this, "Downloading file to your Downloads folder...", Toast.LENGTH_LONG).show()
            } catch (e: Exception) {
                // If direct download request fails (e.g. Blob URL), trigger fallback
                if (url.startsWith("blob:") || url.startsWith("data:")) {
                    Toast.makeText(this, "Download failed: Direct blob downloading not supported natively.", Toast.LENGTH_LONG).show()
                } else {
                    val intent = Intent(Intent.ACTION_VIEW, Uri.parse(url))
                    startActivity(intent)
                }
            }
        }
    }

    private fun checkAndLoadUrl() {
        val sharedPrefs = getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        val savedUrl = sharedPrefs.getString(KEY_SERVER_URL, null)

        if (savedUrl != null) {
            loadServerUrl(savedUrl)
        } else if (DEFAULT_CLOUD_URL.isNotEmpty() && DEFAULT_CLOUD_URL != "https://your-app-name.streamlit.app") {
            saveServerUrl(DEFAULT_CLOUD_URL)
            loadServerUrl(DEFAULT_CLOUD_URL)
        } else {
            showUrlConfigScreen()
        }
    }

    private fun showUrlConfigScreen() {
        webView.visibility = View.GONE
        loadingOverlay.visibility = View.GONE
        errorOverlay.visibility = View.GONE
        urlConfigOverlay.visibility = View.VISIBLE

        val sharedPrefs = getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        val savedUrl = sharedPrefs.getString(KEY_SERVER_URL, "")
        if (savedUrl.isNullOrEmpty()) {
            etServerUrl.setText("https://")
        } else {
            etServerUrl.setText(savedUrl)
        }
        etServerUrl.setSelection(etServerUrl.text.length)
    }

    private fun handleConnect() {
        var inputUrl = etServerUrl.text.toString().trim()
        if (inputUrl.isEmpty() || inputUrl == "https://" || inputUrl == "http://") {
            Toast.makeText(this, "Please enter a valid Streamlit Cloud URL", Toast.LENGTH_SHORT).show()
            return
        }

        // Auto prepend protocol if missing
        if (!inputUrl.startsWith("http://") && !inputUrl.startsWith("https://")) {
            inputUrl = "https://$inputUrl"
        }

        saveServerUrl(inputUrl)
        loadServerUrl(inputUrl)
    }

    private fun saveServerUrl(url: String) {
        val sharedPrefs = getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        sharedPrefs.edit().putString(KEY_SERVER_URL, url).apply()
    }

    private fun loadServerUrl(url: String) {
        urlConfigOverlay.visibility = View.GONE
        errorOverlay.visibility = View.GONE
        loadingOverlay.visibility = View.VISIBLE
        statusText.text = "Connecting to server..."
        webView.loadUrl(url)
    }

    private fun startDiscovery() {
        urlConfigOverlay.visibility = View.GONE
        loadingOverlay.visibility = View.VISIBLE
        errorOverlay.visibility = View.GONE
        webView.visibility = View.GONE
        statusText.text = "Detecting Wi-Fi subnet..."

        val wifiSubnet = getWifiSubnet()
        if (wifiSubnet == null) {
            // No Wi-Fi or local connection, try emulator fallback directly
            statusText.text = "Checking local development emulator link..."
            testAndConnect("http://10.0.2.2:8501")
            return
        }

        statusText.text = "Scanning network on port 8501..."
        
        // Scan subnet asynchronously
        CoroutineScope(Dispatchers.Main).launch {
            val discoveredServerUrl = scanSubnet(wifiSubnet)
            if (discoveredServerUrl != null) {
                statusText.text = "Server discovered! Connecting..."
                webView.loadUrl(discoveredServerUrl)
            } else {
                // Fallback test: Try emulator address
                statusText.text = "Server not found on Wi-Fi. Checking emulator gateway..."
                val emulatorUrl = "http://10.0.2.2:8501"
                val emulatorFound = withContext(Dispatchers.IO) {
                    testPortOpen("10.0.2.2", 8501)
                }
                if (emulatorFound) {
                    statusText.text = "Emulator gateway found! Connecting..."
                    webView.loadUrl(emulatorUrl)
                } else {
                    showErrorState()
                }
            }
        }
    }

    private fun getWifiSubnet(): String? {
        val connectivityManager = getSystemService(Context.CONNECTIVITY_SERVICE) as ConnectivityManager
        val activeNetwork = connectivityManager.activeNetwork ?: return null
        val capabilities = connectivityManager.getNetworkCapabilities(activeNetwork) ?: return null
        
        if (!capabilities.hasTransport(NetworkCapabilities.TRANSPORT_WIFI)) {
            return null
        }

        val wifiManager = applicationContext.getSystemService(Context.WIFI_SERVICE) as WifiManager
        val ipAddress = wifiManager.connectionInfo.ipAddress
        if (ipAddress == 0) return null
        
        // Format to "192.168.1." style base
        return String.format(
            Locale.US,
            "%d.%d.%d.",
            ipAddress and 0xff,
            ipAddress shr 8 and 0xff,
            ipAddress shr 16 and 0xff
        )
    }

    private suspend fun scanSubnet(subnet: String): String? = withContext(Dispatchers.IO) {
        val jobs = mutableListOf<Deferred<String?>>()
        for (i in 1..254) {
            val ip = "$subnet$i"
            val task = async {
                if (testPortOpen(ip, 8501)) {
                    "http://$ip:8501"
                } else {
                    null
                }
            }
            jobs.add(task)
        }
        
        // Return first responding server address
        var foundUrl: String? = null
        for (job in jobs) {
            val url = job.await()
            if (url != null) {
                foundUrl = url
                break
            }
        }
        foundUrl
    }

    private fun testPortOpen(ip: String, port: Int): Boolean {
        return try {
            val socket = Socket()
            socket.connect(InetSocketAddress(ip, port), 250) // 250ms timeout for fast subnet sweeps
            socket.close()
            true
        } catch (e: Exception) {
            false
        }
    }

    private fun testAndConnect(url: String) {
        CoroutineScope(Dispatchers.Main).launch {
            val host = Uri.parse(url).host ?: "10.0.2.2"
            val found = withContext(Dispatchers.IO) {
                testPortOpen(host, 8501)
            }
            if (found) {
                webView.loadUrl(url)
            } else {
                showErrorState()
            }
        }
    }

    private fun showErrorState() {
        loadingOverlay.visibility = View.GONE
        errorOverlay.visibility = View.VISIBLE
        webView.visibility = View.GONE
    }

    override fun onBackPressed() {
        if (webView.canGoBack()) {
            webView.goBack()
        } else {
            super.onBackPressed()
        }
    }
}
