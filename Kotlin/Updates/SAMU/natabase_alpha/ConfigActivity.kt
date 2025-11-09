package com.belasoft.natabase_alpha

import android.app.AlertDialog
import android.content.Intent
import android.os.Bundle
import android.widget.Button
import android.widget.EditText
import android.widget.ImageButton
import android.widget.Switch
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import com.belasoft.natabase_alpha.utils.SettingsManager
import java.io.File

class ConfigActivity : BaseActivity() {

    private lateinit var btnAlterarDiretorio: Button
    private lateinit var btnLimparCache: Button
    private lateinit var btnSalvarEmailDestino: Button
    private lateinit var switchEmailAuto: Switch
    private lateinit var textViewDiretorioAtual: TextView
    private lateinit var textViewEmailAtual: TextView
    private lateinit var editTextEmailDestino: EditText

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_config)

        setupDrawer()

        initViews()
        setupClickListeners()
        loadCurrentSettings()
    }

    private fun initViews() {
        btnAlterarDiretorio = findViewById(R.id.btnAlterarDiretorio)
        btnLimparCache = findViewById(R.id.btnLimparCache)
        btnSalvarEmailDestino = findViewById(R.id.btnSalvarEmailDestino)
        switchEmailAuto = findViewById(R.id.switchEmailAuto)
        textViewDiretorioAtual = findViewById(R.id.textViewDiretorioAtual)
        textViewEmailAtual = findViewById(R.id.textViewEmailAtual)
        editTextEmailDestino = findViewById(R.id.editTextEmailDestino)
    }

    private fun setupClickListeners() {
        btnAlterarDiretorio.setOnClickListener {
            showDirectorySelectionDialog()
        }

        btnLimparCache.setOnClickListener {
            showClearCacheConfirmation()
        }

        btnSalvarEmailDestino.setOnClickListener {
            saveDestinationEmail()
        }

        switchEmailAuto.setOnCheckedChangeListener { _, isChecked ->
            SettingsManager.setAutoEmailEnabled(this, isChecked)
            Toast.makeText(this,
                if (isChecked) "Email automático ativado" else "Email automático desativado",
                Toast.LENGTH_SHORT).show()
        }
    }

    private fun loadCurrentSettings() {
        switchEmailAuto.isChecked = SettingsManager.isAutoEmailEnabled(this)

        // Carregar email de destino
        val destinationEmail = SettingsManager.getDestinationEmail(this)
        if (destinationEmail.isNotBlank()) {
            textViewEmailAtual.text = "Email atual: $destinationEmail"
            editTextEmailDestino.setText(destinationEmail)
        } else {
            textViewEmailAtual.text = "Email atual: Não definido"
        }

        updateDirectoryDisplay()
    }

    private fun saveDestinationEmail() {
        val email = editTextEmailDestino.text.toString().trim()

        if (email.isEmpty()) {
            Toast.makeText(this, "Por favor, digite um email", Toast.LENGTH_SHORT).show()
            return
        }

        if (!isValidEmail(email)) {
            Toast.makeText(this, "Por favor, digite um email válido", Toast.LENGTH_SHORT).show()
            return
        }

        SettingsManager.setDestinationEmail(this, email)
        textViewEmailAtual.text = "Email atual: $email"
        editTextEmailDestino.setText("")

        Toast.makeText(this, "Email de destino salvo com sucesso", Toast.LENGTH_SHORT).show()
    }

    private fun isValidEmail(email: String): Boolean {
        return android.util.Patterns.EMAIL_ADDRESS.matcher(email).matches()
    }

    private fun updateDirectoryDisplay() {
        val currentDir = SettingsManager.getExportDirectory(this)
        val directoryFile = SettingsManager.getExportDirectoryFile(this)

        textViewDiretorioAtual.text = when (currentDir) {
            SettingsManager.KEY_DEFAULT_EXPORT_DIR -> "Diretório da aplicação\n(${directoryFile.absolutePath})"
            SettingsManager.KEY_DOCUMENTS_DIR -> "Documentos públicos\n(${directoryFile.absolutePath})"
            SettingsManager.KEY_DOWNLOADS_DIR -> "Downloads públicos\n(${directoryFile.absolutePath})"
            else -> "Diretório da aplicação\n(${directoryFile.absolutePath})"
        }

        if (!SettingsManager.isDirectoryAvailable(this)) {
            textViewDiretorioAtual.text = "${textViewDiretorioAtual.text}\nDiretório não disponível"
        }
    }

    private fun showDirectorySelectionDialog() {
        val directories = arrayOf(
            "Diretório da aplicação (Recomendado)",
            "Documentos públicos",
            "Downloads públicos"
        )

        AlertDialog.Builder(this)
            .setTitle("Selecionar Diretório de Exportação")
            .setItems(directories) { _, which ->
                when (which) {
                    0 -> setExportDirectory(SettingsManager.KEY_DEFAULT_EXPORT_DIR)
                    1 -> setExportDirectory(SettingsManager.KEY_DOCUMENTS_DIR)
                    2 -> setExportDirectory(SettingsManager.KEY_DOWNLOADS_DIR)
                }
            }
            .setNegativeButton("Cancelar", null)
            .show()
    }

    private fun setExportDirectory(directory: String) {
        SettingsManager.setExportDirectory(this, directory)

        val directoryFile = SettingsManager.getExportDirectoryFile(this)
        if (!directoryFile.exists()) {
            directoryFile.mkdirs()
        }

        updateDirectoryDisplay()

        val message = when (directory) {
            SettingsManager.KEY_DEFAULT_EXPORT_DIR -> "Diretório alterado para o diretório da aplicação"
            SettingsManager.KEY_DOCUMENTS_DIR -> {
                if (SettingsManager.isDirectoryAvailable(this)) {
                    "Diretório alterado para Documentos públicos"
                } else {
                    "Diretório de Documentos pode não estar disponível. A usar diretório alternativo."
                }
            }
            SettingsManager.KEY_DOWNLOADS_DIR -> {
                if (SettingsManager.isDirectoryAvailable(this)) {
                    "Diretório alterado para Downloads públicos"
                } else {
                    "Diretório de Downloads pode não estar disponível. A usar diretório alternativo."
                }
            }
            else -> "Diretório alterado"
        }

        Toast.makeText(this, message, Toast.LENGTH_LONG).show()
    }

    private fun showClearCacheConfirmation() {
        AlertDialog.Builder(this)
            .setTitle("Limpar Cache")
            .setMessage("Tem a certeza que deseja limpar todos os dados temporários da aplicação? Esta ação não pode ser desfeita.")
            .setPositiveButton("Limpar") { _, _ ->
                clearCache()
            }
            .setNegativeButton("Cancelar", null)
            .show()
    }

    private fun clearCache() {
        SettingsManager.clearCacheData(this)
        Toast.makeText(this, "Cache limpo com sucesso", Toast.LENGTH_SHORT).show()
        loadCurrentSettings()
    }

    override fun onResume() {
        super.onResume()
        loadCurrentSettings()
    }
}