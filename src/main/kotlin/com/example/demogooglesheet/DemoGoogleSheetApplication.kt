package com.example.demogooglesheet

import com.google.api.client.auth.oauth2.Credential
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport
import com.google.api.client.http.HttpExecuteInterceptor
import com.google.api.client.http.HttpRequest
import com.google.api.client.http.HttpRequestInitializer
import com.google.api.client.http.HttpTransport
import com.google.api.client.http.apache.v2.ApacheHttpTransport
import com.google.api.client.http.javanet.NetHttpTransport
import com.google.api.client.json.JsonFactory
import com.google.api.client.json.jackson2.JacksonFactory
import com.google.api.client.util.store.FileDataStoreFactory
import com.google.api.client.util.store.MemoryDataStoreFactory
import com.google.api.services.sheets.v4.Sheets
import com.google.api.services.sheets.v4.SheetsScopes
import com.google.api.services.sheets.v4.model.Spreadsheet
import com.google.api.services.sheets.v4.model.SpreadsheetProperties
import com.google.api.services.sheets.v4.model.ValueRange
import com.google.auth.http.HttpCredentialsAdapter
import com.google.auth.http.HttpTransportFactory
import com.google.auth.oauth2.GoogleCredentials
import com.google.auth.oauth2.ServiceAccountCredentials
import org.apache.http.HttpHost
import org.apache.http.auth.AuthScope
import org.apache.http.auth.UsernamePasswordCredentials
import org.apache.http.client.CredentialsProvider
import org.apache.http.client.HttpClient
import org.apache.http.conn.routing.HttpRoutePlanner
import org.apache.http.impl.client.BasicCredentialsProvider
import org.apache.http.impl.client.ProxyAuthenticationStrategy
import org.apache.http.impl.conn.DefaultProxyRoutePlanner
import org.slf4j.LoggerFactory
import org.springframework.beans.factory.annotation.Value
import org.springframework.boot.CommandLineRunner
import org.springframework.boot.autoconfigure.SpringBootApplication
import org.springframework.boot.runApplication
import java.io.*


@SpringBootApplication
class DemoGoogleSheetApplication: CommandLineRunner {

	val HTTP_TRANSPORT: NetHttpTransport = GoogleNetHttpTransport.newTrustedTransport()

	@Value("\${demo.proxyHost:#{null}}")
	var proxyHost: String? = null

	@Value("\${demo.proxyPort:#{null}}")
	var proxyPort: Int? = null

	@Value("\${demo.proxyUsername:#{null}}")
	var proxyUsername: String? = null

	@Value("\${demo.proxyPassword:#{null}}")
	var proxyPassword: String? = null

	@Value("\${demo.serviceAccountResourceJson}")
	private lateinit var serviceAccountResourceJson: String

	@Value("\${demo.sheetId}")
	private lateinit var sheetId: String

	val httpTransportFactory = getHttpTransportFactory(proxyHost, proxyPort, proxyUsername, proxyPassword)
	val JSON_FACTORY: JsonFactory = JacksonFactory.getDefaultInstance()

	companion object {
		val log = LoggerFactory.getLogger(DemoGoogleSheetApplication::class.java)
	}

	fun getHttpTransportFactory(
		proxyHost: String? = null,
		proxyPort: Int? = null,
		proxyUsername: String? = null,
		proxyPassword: String? = null
	): HttpTransportFactory {

		val httpClient: HttpClient = ApacheHttpTransport.newDefaultHttpClientBuilder().apply {

			if (proxyHost != null && proxyPort != null) {
				val proxyHostDetails = HttpHost(proxyHost, proxyPort)
				val httpRoutePlanner: HttpRoutePlanner = DefaultProxyRoutePlanner(proxyHostDetails)

				setRoutePlanner(httpRoutePlanner)
				setProxyAuthenticationStrategy(ProxyAuthenticationStrategy.INSTANCE)

				if (proxyUsername != null && proxyPassword != null) {
					val credentialsProvider: CredentialsProvider = BasicCredentialsProvider().apply {
						setCredentials(
							AuthScope(proxyHostDetails.hostName, proxyHostDetails.port),
							UsernamePasswordCredentials(proxyUsername, proxyPassword)
						)
					}

					setDefaultCredentialsProvider(credentialsProvider)
				}
			}

		}.build()

		val httpTransport: HttpTransport = ApacheHttpTransport(httpClient)
		return HttpTransportFactory { httpTransport }
	}

	@Throws(IOException::class)
	private fun getCredentials(HTTP_TRANSPORT: NetHttpTransport): Credential? {
		// Load client secrets.
		val ins: InputStream = this::class.java.getResourceAsStream("/client_secret_400300917728-ifqj4hglg5l4vbdr5ck0nvof6dsi0ini.apps.googleusercontent.com.json")
			?: throw FileNotFoundException("Resource not found: /client_secret_400300917728-ifqj4hglg5l4vbdr5ck0nvof6dsi0ini.apps.googleusercontent.com.json")
		val clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, InputStreamReader(ins))

		// Build flow and trigger user authorization request.
		val flow = GoogleAuthorizationCodeFlow.Builder(
			HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, listOf(SheetsScopes.SPREADSHEETS)
		)
			.setDataStoreFactory(MemoryDataStoreFactory())
			.setAccessType("offline")
			.build()
		val receiver = LocalServerReceiver.Builder().setPort(8888).build()
		return AuthorizationCodeInstalledApp(flow, receiver).authorize("user")
	}

	fun getServiceAccountCredential(scopes: List<String>): HttpRequestInitializer {
		/* deprecated method */
		// return GoogleCredential.fromStream(ins).createScoped(listOf(SheetsScopes.SPREADSHEETS))
		val ins: InputStream = this::class.java.getResourceAsStream(serviceAccountResourceJson)
		return HttpCredentialsAdapter(ServiceAccountCredentials.fromStream(ins, httpTransportFactory).createScoped(scopes))
	}

	override fun run(vararg args: String?) {
		val dummyProxyPassword = if (proxyPassword == null) null else "***"
		log.info("proxyHost: $proxyHost, proxyPort: $proxyPort, proxyUsername: $proxyUsername, proxyPassword: ${dummyProxyPassword}")
//		val service = Sheets.Builder(HTTP_TRANSPORT,
//			JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
//			.setApplicationName("Google Sheets API Java Quickstart")
//			.build()

		val service = Sheets.Builder(
			httpTransportFactory.create(),
			JSON_FACTORY,
			getServiceAccountCredential(listOf(SheetsScopes.SPREADSHEETS))
		).setApplicationName("Google Sheets API Java Quickstart")
			.build()

		val spreadsheetId = sheetId

		for (i in 1..10) {
			val values = listOf(
				listOf("A $i", "B $i", "C $i")
			)
			val body: ValueRange = ValueRange()
				.setValues(values)
			val result = service.spreadsheets().values().append(spreadsheetId, "Sheet1!A2", body)
				.setValueInputOption("RAW")
				.execute()
		}
	}
}

fun main(args: Array<String>) {
	runApplication<DemoGoogleSheetApplication>(*args)
}