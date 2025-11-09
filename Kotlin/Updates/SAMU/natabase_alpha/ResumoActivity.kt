package com.belasoft.natabase_alpha

import android.app.AlertDialog
import android.content.Context
import android.content.Intent
import android.graphics.Typeface
import android.net.ConnectivityManager
import android.net.NetworkCapabilities
import android.os.Bundle
import android.view.View
import android.widget.*
import androidx.appcompat.app.AppCompatActivity
import androidx.core.content.ContextCompat
import androidx.core.view.GravityCompat
import androidx.drawerlayout.widget.DrawerLayout
import androidx.lifecycle.lifecycleScope
import com.belasoft.natabase_alpha.utils.CacheManager
import com.belasoft.natabase_alpha.utils.SettingsManager
import com.google.android.material.navigation.NavigationView
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import java.io.File
import java.net.InetSocketAddress
import java.net.Socket
import java.text.SimpleDateFormat
import java.util.*

class ResumoActivity : BaseActivity() {

    private lateinit var tableContainer: LinearLayout
    private lateinit var btnEnviarEmail: Button
    private lateinit var btnVoltar: ImageButton
    private lateinit var mapaProducao: MapaProducao
    private var producaoAtualIndex: Int = 1

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_resumo)

        setupDrawer()

        initViews()
        setupData()
        mostrarResumo()

        btnEnviarEmail.setOnClickListener {
            val directory = SettingsManager.getExportDirectoryFile(this)
            val file = File(directory, "mapa_producao.xlsx")
            if (file.exists()) {
                enviarEmail(file)
            } else {
                Toast.makeText(this, "Arquivo não encontrado", Toast.LENGTH_SHORT).show()
            }
        }

        btnVoltar.setOnClickListener {
            finish()
        }
    }

    private fun initViews() {
        tableContainer = findViewById(R.id.tableContainer)
        btnEnviarEmail = findViewById(R.id.btnEnviarEmail)
        btnVoltar = findViewById(R.id.btnVoltar)
    }

    private fun setupData() {
        mapaProducao = CacheManager.carregarProducao(this)
            ?: ExcelService.carregarMapaProducao(this)

        producaoAtualIndex = CacheManager.carregarProducaoIndex(this)
    }

    private fun mostrarResumo() {
        tableContainer.removeAllViews()

        // ATUALIZAÇÃO: Incluir header com data e informações gerais
        val headerContainer = LinearLayout(this).apply {
            orientation = LinearLayout.VERTICAL
            background = ContextCompat.getDrawable(this@ResumoActivity, R.drawable.rounded_button_brown)
            setPadding(24, 16, 24, 16)
            val params = LinearLayout.LayoutParams(
                LinearLayout.LayoutParams.MATCH_PARENT,
                LinearLayout.LayoutParams.WRAP_CONTENT
            )
            params.setMargins(0, 0, 0, 20)
            layoutParams = params
        }

        val dataAtual = if (mapaProducao.data.isNotBlank()) {
            mapaProducao.data
        } else {
            SimpleDateFormat("dd/MM/yyyy", Locale.getDefault()).format(Date())
        }

        val diaSemana = if (mapaProducao.diaSemana.isNotBlank()) {
            mapaProducao.diaSemana
        } else {
            SimpleDateFormat("EEEE", Locale("pt", "PT")).format(Date()).capitalize(Locale.getDefault())
        }

        val tvData = TextView(this).apply {
            text = "$diaSemana - $dataAtual"
            textSize = 20f
            setTextColor(ContextCompat.getColor(this@ResumoActivity, R.color.white))
            setTypeface(null, Typeface.BOLD)
            gravity = android.view.Gravity.CENTER
        }

        val tvTitulo = TextView(this).apply {
            text = "Resumo de Produção"
            textSize = 24f
            setTextColor(ContextCompat.getColor(this@ResumoActivity, R.color.white))
            setTypeface(null, Typeface.BOLD)
            gravity = android.view.Gravity.CENTER
            setPadding(0, 8, 0, 0)
        }

        headerContainer.addView(tvData)
        headerContainer.addView(tvTitulo)
        tableContainer.addView(headerContainer)

        val itensPorCategoria = mapaProducao.itens
            .filter { item ->
                item.producoes.isNotEmpty() || item.perdas > 0 || item.sobras > 0
            }
            .groupBy { it.categoria }

        for ((categoria, itens) in itensPorCategoria) {
            val categoriaContainer = LinearLayout(this).apply {
                orientation = LinearLayout.VERTICAL
                background = ContextCompat.getDrawable(this@ResumoActivity, R.drawable.rounded_button_brown)
                setPadding(0, 0, 0, 0)
                val params = LinearLayout.LayoutParams(
                    LinearLayout.LayoutParams.MATCH_PARENT,
                    LinearLayout.LayoutParams.WRAP_CONTENT
                )
                params.setMargins(0, 20, 0, 0)
                layoutParams = params
            }

            val header = TextView(this).apply {
                text = categoria
                textSize = 22f
                setTextColor(ContextCompat.getColor(this@ResumoActivity, R.color.caramel))
                setTypeface(null, Typeface.BOLD)
                gravity = android.view.Gravity.CENTER
                setPadding(24, 16, 24, 16)
            }

            val linha = View(this).apply {
                layoutParams = LinearLayout.LayoutParams(
                    LinearLayout.LayoutParams.MATCH_PARENT,
                    2
                )
                setBackgroundColor(ContextCompat.getColor(this@ResumoActivity, R.color.caramel))
            }

            val itensContainer = LinearLayout(this).apply {
                orientation = LinearLayout.VERTICAL
                setPadding(24, 16, 24, 16)
                setBackgroundColor(ContextCompat.getColor(this@ResumoActivity, R.color.caramel))
            }

            categoriaContainer.addView(header)
            categoriaContainer.addView(linha)
            categoriaContainer.addView(itensContainer)

            itens.forEach { item ->
                val producoesComQuantidade = item.producoes.filter { it.quantidade > 0 }
                val total = item.producoes.sumOf { it.quantidade }

                if (total > 0 || item.perdas > 0 || item.sobras > 0) {
                    val produtoContainer = LinearLayout(this).apply {
                        orientation = LinearLayout.VERTICAL
                        setPadding(0, 8, 0, 16)
                    }

                    // Nome do produto e total
                    val produtoHeader = TextView(this).apply {
                        text = "${item.produto} - Total: $total"
                        textSize = 18f
                        setTextColor(ContextCompat.getColor(this@ResumoActivity, R.color.white))
                        setTypeface(null, Typeface.BOLD)
                    }
                    produtoContainer.addView(produtoHeader)

                    producoesComQuantidade.forEachIndexed { index, producao ->
                        if (producao.quantidade > 0) {
                            val numeroProducao = index + 1
                            val detalheProducao = TextView(this).apply {
                                text = "  Produção $numeroProducao: ${producao.quantidade} unidades (${producao.hora})"
                                textSize = 14f
                                setTextColor(ContextCompat.getColor(this@ResumoActivity, R.color.white))
                                setPadding(16, 2, 0, 2)
                            }
                            produtoContainer.addView(detalheProducao)
                        }
                    }

                    // Perdas e sobras
                    if (item.perdas > 0) {
                        val perdasText = TextView(this).apply {
                            text = "  Perdas: ${item.perdas} unidades"
                            textSize = 14f
                            setTextColor(ContextCompat.getColor(this@ResumoActivity, R.color.red))
                            setPadding(16, 2, 0, 2)
                        }
                        produtoContainer.addView(perdasText)
                    }

                    if (item.sobras > 0) {
                        val sobrasText = TextView(this).apply {
                            text = "  Sobras: ${item.sobras} unidades"
                            textSize = 14f
                            setTextColor(ContextCompat.getColor(this@ResumoActivity, R.color.green))
                            setPadding(16, 2, 0, 2)
                        }
                        produtoContainer.addView(sobrasText)
                    }

                    // Saldo final
                    val saldo = total - (item.perdas + item.sobras)
                    val saldoText = TextView(this).apply {
                        text = "Total: $saldo unidades"
                        textSize = 16f
                        setTextColor(ContextCompat.getColor(this@ResumoActivity,
                            if (saldo >= 0) R.color.white else R.color.red))
                        setTypeface(null, Typeface.BOLD)
                        setPadding(16, 8, 0, 0)
                    }
                    produtoContainer.addView(saldoText)

                    itensContainer.addView(produtoContainer)
                }
            }

            tableContainer.addView(categoriaContainer)

            val espaco = View(this).apply {
                layoutParams = LinearLayout.LayoutParams(
                    LinearLayout.LayoutParams.MATCH_PARENT,
                    20
                )
            }
            tableContainer.addView(espaco)
        }

        if (itensPorCategoria.isEmpty()) {
            val tvVazio = TextView(this).apply {
                text = "Nenhuma produção registada"
                textSize = 18f
                setTextColor(ContextCompat.getColor(this@ResumoActivity, R.color.deep_brown))
                setTypeface(null, Typeface.BOLD)
                gravity = android.view.Gravity.CENTER
                setPadding(0, 32, 0, 32)
            }
            tableContainer.addView(tvVazio)
        }
    }

    private suspend fun testarConectividade(): Boolean = withContext(Dispatchers.IO) {
        try {
            val timeout = 10000
            Socket().use { socket ->
                val socketAddress = InetSocketAddress("smtp.gmail.com", 587)
                socket.connect(socketAddress, timeout)
            }
            true
        } catch (e: Exception) {
            e.printStackTrace()
            false
        }
    }

    private fun enviarEmail(file: File) {
        AlertDialog.Builder(this)
            .setTitle("Adicionar Perdas e Sobras")
            .setMessage("Deseja adicionar perdas e sobras antes de enviar o email?")
            .setPositiveButton("Sim") { _, _ ->
                val intent = Intent(this, PerdasSobrasActivity::class.java).apply {
                    putExtra("MAPA_PRODUCAO", mapaProducao)
                    putExtra("PRODUCAO_INDEX", producaoAtualIndex)
                }
                startActivity(intent)
            }
            .setNegativeButton("Não, Enviar Agora") { _, _ ->
                enviarEmailDiretamente(file)
            }
            .setNeutralButton("Cancelar", null)
            .show()
    }

    private fun enviarEmailDiretamente(file: File) {
        val emailDestinatario = EmailService.getDestinationEmail(this)

        val dataParaEmail = if (mapaProducao.data.isNotBlank()) {
            mapaProducao.data
        } else {
            val hoje = java.text.SimpleDateFormat("dd/MM/yyyy", java.util.Locale.getDefault()).format(java.util.Date())
            hoje
        }

        lifecycleScope.launch {
            val podeConectar = testarConectividade()
            if (!podeConectar) {
                Toast.makeText(
                    this@ResumoActivity,
                    "Não foi possível conectar ao servidor de email",
                    Toast.LENGTH_LONG
                ).show()
                return@launch
            }

            withContext(Dispatchers.IO) {
                try {
                    val emailService = EmailService(
                        senderEmail = "relatorioloja012@gmail.com",
                        senderAppPassword = "cwvt qgcg etrd ydzw",
                        toAddresses = listOf(emailDestinatario)
                    )
                    emailService.sendExcel(
                        file,
                        "Mapa de Produção - $dataParaEmail",
                        "Segue o mapa de produção em anexo."
                    )
                    withContext(Dispatchers.Main) {
                        Toast.makeText(
                            this@ResumoActivity,
                            "Email enviado com sucesso para $emailDestinatario!",
                            Toast.LENGTH_SHORT
                        ).show()
                    }
                } catch (e: Exception) {
                    e.printStackTrace()
                    withContext(Dispatchers.Main) {
                        when (e) {
                            is java.net.UnknownHostException ->
                                Toast.makeText(
                                    this@ResumoActivity,
                                    "Erro de DNS: Não foi possível encontrar o servidor de email",
                                    Toast.LENGTH_LONG
                                ).show()

                            is java.net.ConnectException ->
                                Toast.makeText(
                                    this@ResumoActivity,
                                    "Erro de conexão: Verifique sua internet",
                                    Toast.LENGTH_LONG
                                ).show()

                            else ->
                                Toast.makeText(
                                    this@ResumoActivity,
                                    "Erro ao enviar email: ${e.message}",
                                    Toast.LENGTH_LONG
                                ).show()
                        }
                    }
                }
            }
        }
    }

    private fun isNetworkAvailable(context: Context): Boolean {
        val cm = context.getSystemService(Context.CONNECTIVITY_SERVICE) as ConnectivityManager
        val nw = cm.activeNetwork ?: return false
        val actNw = cm.getNetworkCapabilities(nw) ?: return false
        return actNw.hasCapability(NetworkCapabilities.NET_CAPABILITY_INTERNET)
    }
}