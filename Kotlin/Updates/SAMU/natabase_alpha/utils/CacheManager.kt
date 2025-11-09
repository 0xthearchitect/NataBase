package com.belasoft.natabase_alpha.utils

import android.content.Context
import com.belasoft.natabase_alpha.MapaProducao
import com.google.gson.Gson
import com.google.gson.reflect.TypeToken

object CacheManager {
    private const val PREFS_NAME = "ProducaoPrefs"
    private const val KEY_MAPA = "mapa_producao"
    private const val KEY_DATA = "data"
    private const val KEY_PRODUCAO_INDEX = "producao_index"

    fun salvarProducao(context: Context, mapa: MapaProducao) {
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        val editor = prefs.edit()

        val json = Gson().toJson(mapa)
        editor.putString(KEY_MAPA, json)

        val hoje = java.text.SimpleDateFormat("dd/MM/yyyy", java.util.Locale.getDefault()).format(java.util.Date())
        editor.putString(KEY_DATA, hoje)

        editor.apply()
    }

    fun carregarProducao(context: Context): MapaProducao? {
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)

        val hoje = java.text.SimpleDateFormat("dd/MM/yyyy", java.util.Locale.getDefault()).format(java.util.Date())
        val dataSalva = prefs.getString(KEY_DATA, "")

        return if (dataSalva == hoje) {
            val json = prefs.getString(KEY_MAPA, null)
            if (json != null) {
                val type = object : TypeToken<MapaProducao>() {}.type
                Gson().fromJson(json, type)
            } else {
                null
            }
        } else {
            prefs.edit().remove(KEY_MAPA).remove(KEY_PRODUCAO_INDEX).apply()
            null
        }
    }

    fun salvarProducaoIndex(context: Context, index: Int) {
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        prefs.edit().putInt(KEY_PRODUCAO_INDEX, index).apply()
    }

    fun carregarProducaoIndex(context: Context): Int {
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)

        val hoje = java.text.SimpleDateFormat("dd/MM/yyyy", java.util.Locale.getDefault()).format(java.util.Date())
        val dataSalva = prefs.getString(KEY_DATA, "")

        return if (dataSalva == hoje) {
            prefs.getInt(KEY_PRODUCAO_INDEX, 1)
        } else {
            prefs.edit().putInt(KEY_PRODUCAO_INDEX, 1).apply()
            1
        }
    }

    fun clearAllData(context: Context) {
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        prefs.edit().clear().apply()
    }
}