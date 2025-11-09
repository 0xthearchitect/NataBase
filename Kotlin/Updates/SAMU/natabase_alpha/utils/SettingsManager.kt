package com.belasoft.natabase_alpha.utils

import android.content.Context
import android.content.SharedPreferences
import android.os.Environment
import java.io.File

object SettingsManager {
    private const val PREFS_NAME = "AppSettings"
    private const val KEY_EMAIL_AUTO = "email_auto_enabled"
    private const val KEY_EXPORT_DIRECTORY = "export_directory"

    const val KEY_DEFAULT_EXPORT_DIR = "default"
    const val KEY_DOCUMENTS_DIR = "documents"
    const val KEY_DOWNLOADS_DIR = "downloads"
    private const val KEY_DESTINATION_EMAIL = "exemplo@gmail.com"

    fun isAutoEmailEnabled(context: Context): Boolean {
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        return prefs.getBoolean(KEY_EMAIL_AUTO, true)
    }

    fun setAutoEmailEnabled(context: Context, enabled: Boolean) {
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        prefs.edit().putBoolean(KEY_EMAIL_AUTO, enabled).apply()
    }

    fun getExportDirectory(context: Context): String {
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        return prefs.getString(KEY_EXPORT_DIRECTORY, KEY_DEFAULT_EXPORT_DIR) ?: KEY_DEFAULT_EXPORT_DIR
    }

    fun setExportDirectory(context: Context, directory: String) {
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        prefs.edit().putString(KEY_EXPORT_DIRECTORY, directory).apply()
    }

    fun getExportDirectoryFile(context: Context): File {
        return when (getExportDirectory(context)) {
            KEY_DOCUMENTS_DIR -> {
                Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS)
                    ?: context.getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS)
                    ?: context.filesDir
            }
            KEY_DOWNLOADS_DIR -> {
                Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS)
                    ?: context.getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS)
                    ?: context.filesDir
            }
            else -> {
                context.getExternalFilesDir(null) ?: context.filesDir
            }
        }
    }

    fun clearAllSettings(context: Context) {
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        prefs.edit().clear().apply()
    }

    fun clearCacheData(context: Context) {
        CacheManager.clearAllData(context)

        try {
            val cacheDir = context.cacheDir
            if (cacheDir.exists() && cacheDir.isDirectory) {
                cacheDir.listFiles()?.forEach { file ->
                    file.delete()
                }
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }

    fun getDestinationEmail(context: Context): String {
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        return prefs.getString(KEY_DESTINATION_EMAIL, "") ?: ""
    }

    fun setDestinationEmail(context: Context, email: String) {
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        prefs.edit().putString(KEY_DESTINATION_EMAIL, email).apply()
    }

    fun hasDestinationEmail(context: Context): Boolean {
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        return prefs.getString(KEY_DESTINATION_EMAIL, "")?.isNotBlank() ?: false
    }

    fun isDirectoryAvailable(context: Context): Boolean {
        return try {
            val dir = getExportDirectoryFile(context)
            dir.exists() || dir.mkdirs()
        } catch (e: Exception) {
            false
        }
    }
}