package com.belasoft.natabase_alpha

import android.content.Context
import com.belasoft.natabase_alpha.utils.SettingsManager
import java.io.File
import java.util.*
import javax.mail.*
import javax.mail.internet.InternetAddress
import javax.mail.internet.MimeBodyPart
import javax.mail.internet.MimeMessage
import javax.mail.internet.MimeMultipart

class EmailService(
    private val senderEmail: String,
    private val senderAppPassword: String,
    private val toAddresses: List<String>
) {
    private val smtpHost = "smtp.gmail.com"
    private val smtpPort = "587"

    fun sendExcel(file: File, subject: String, body: String) {
        val props = Properties().apply {
            put("mail.smtp.host", smtpHost)
            put("mail.smtp.port", smtpPort)
            put("mail.smtp.auth", "true")
            put("mail.smtp.starttls.enable", "true")
            put("mail.smtp.ssl.trust", smtpHost)
            put("mail.smtp.ssl.protocols", "TLSv1.2 TLSv1.3")
            put("mail.smtp.connectiontimeout", "30000")
            put("mail.smtp.timeout", "30000")
            put("mail.smtp.writetimeout", "30000")
        }

        val session = Session.getInstance(props,
            object : Authenticator() {
                override fun getPasswordAuthentication(): PasswordAuthentication {
                    return PasswordAuthentication(senderEmail, senderAppPassword)
                }
            })

        try {
            val message = MimeMessage(session)
            message.setFrom(InternetAddress(senderEmail))

            val addresses = toAddresses.map { InternetAddress(it) }.toTypedArray()
            message.setRecipients(Message.RecipientType.TO, addresses)

            message.subject = subject

            val multipart = MimeMultipart()

            val textPart = MimeBodyPart()
            textPart.setText(body)
            multipart.addBodyPart(textPart)

            val attachPart = MimeBodyPart()
            attachPart.attachFile(file)
            multipart.addBodyPart(attachPart)

            message.setContent(multipart)

            Transport.send(message)
            println("Email enviado com sucesso.")

        } catch (e: Exception) {
            e.printStackTrace()
            throw RuntimeException("Erro ao enviar email: ${e.message}")
        }
    }

    companion object {
        fun shouldSendAutoEmail(context: Context): Boolean {
            return SettingsManager.isAutoEmailEnabled(context)
        }

        fun getDestinationEmail(context: Context): String {
            val customEmail = SettingsManager.getDestinationEmail(context)
            return if (customEmail.isNotBlank()) {
                customEmail
            } else {
                // Email padrão caso não esteja configurado
                "exemplo@gmail.com"
            }
        }
    }
}