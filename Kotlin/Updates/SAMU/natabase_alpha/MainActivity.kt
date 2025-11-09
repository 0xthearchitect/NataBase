package com.belasoft.natabase_alpha

import android.app.AlertDialog
import android.content.Intent
import android.content.res.Configuration
import android.os.Bundle
import android.view.View
import android.view.animation.AnimationUtils
import android.widget.Button
import android.widget.FrameLayout
import android.widget.ImageButton
import android.widget.LinearLayout
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import androidx.core.content.ContextCompat
import androidx.core.view.GravityCompat
import androidx.drawerlayout.widget.DrawerLayout
import androidx.recyclerview.widget.GridLayoutManager
import androidx.recyclerview.widget.RecyclerView
import com.belasoft.natabase_alpha.utils.CacheManager
import com.google.android.material.navigation.NavigationView
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

class MainActivity : BaseActivity(), CalculadoraFragment.CalculadoraListener {

    private lateinit var mapaProducao: MapaProducao
    private lateinit var recycler: RecyclerView
    private var producaoAtualIndex = 1
    private lateinit var currentAdapter: ItemProducaoAdapter
    private var categoriaAtual: String = ""
    private var isTabletMode: Boolean = false
    private var calculadoraAberta: Boolean = false

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_home)

        // Configurar o drawer
        setupDrawer()

        inicializarMapaProducao()
        atualizarMensagemBoasVindas()
        configurarBotoesHome()
        adicionarBotaoDebug()

        // Verificar se deve abrir inventário diretamente
        if (intent.getBooleanExtra("OPEN_INVENTORY", false)) {
            abrirInventario()
        }

        // NOVO: Verificar se deve abrir produção diretamente
        if (intent.getBooleanExtra("OPEN_PRODUCTION", false)) {
            abrirProducao()
        }
    }

    fun abrirProducao() {
        setContentView(R.layout.activity_producao)
        setupDrawer() // Reconfigurar o drawer para este layout

        inicializarMapaProducao()
        recycler = findViewById(R.id.recyclerProdutos)

        // Configurar botão voltar específico
        findViewById<ImageButton>(R.id.btnVoltar)?.setOnClickListener {
            voltarParaHome()
        }

        adicionarBotaoDebug()
        carregarCategorias()
    }

    private fun mostrarDialogoSair() {
        AlertDialog.Builder(this)
            .setTitle("Sair do Aplicativo")
            .setMessage("Deseja realmente sair do aplicativo?")
            .setPositiveButton("Sim") { _, _ ->
                finishAffinity() // Fecha completamente o aplicativo
            }
            .setNegativeButton("Cancelar", null)
            .show()
    }

    private fun inicializarMapaProducao() {
        val hoje = SimpleDateFormat("dd/MM/yyyy", Locale.getDefault()).format(Date())
        val dataCache = CacheManager.carregarProducao(this)?.data ?: ""

        if (dataCache != hoje) {
            CacheManager.salvarProducaoIndex(this, 1)
        }

        producaoAtualIndex = CacheManager.carregarProducaoIndex(this)
        val mapaCache = CacheManager.carregarProducao(this)
        mapaProducao = (mapaCache ?: ExcelService.carregarMapaProducao(this))
    }

    private fun abrirProdutosDaCategoria(categoria: String) {
        try {
            categoriaAtual = categoria
            val todosProdutosCategoria = mapaProducao.itens.filter { it.categoria == categoria }
            val produtosCategoria = todosProdutosCategoria.toMutableList()

            val layoutManager = GridLayoutManager(this, calcularNumeroColunasProdutos())
            recycler.layoutManager = layoutManager

            findViewById<ImageButton>(R.id.btnVoltar)?.setOnClickListener {
                voltarParaCategorias()
            }

            val statusText = findViewById<TextView>(R.id.statusText)
            val dataAtual = SimpleDateFormat("dd/MM/yyyy", Locale.getDefault()).format(Date())
            statusText?.text = "$categoria - $dataAtual"

            currentAdapter = ItemProducaoAdapter(
                produtosCategoria,
                onQuantidadeAlterada = { item ->
                    try {
                        val index = mapaProducao.itens.indexOfFirst { it.produto == item.produto && it.categoria == item.categoria }
                        if (index != -1) {
                            val novaLista = mapaProducao.itens.toMutableList()
                            novaLista[index] = item
                            mapaProducao = mapaProducao.copy(itens = novaLista)
                            CacheManager.salvarProducao(this, mapaProducao)
                        }
                    } catch (e: Exception) {
                        e.printStackTrace()
                    }
                },
                salvarCallback = { itensAtualizados ->
                    try {
                        itensAtualizados.forEach { itemAtualizado ->
                            val index = mapaProducao.itens.indexOfFirst {
                                it.produto == itemAtualizado.produto && it.categoria == itemAtualizado.categoria
                            }
                            if (index != -1) {
                                val novaLista = mapaProducao.itens.toMutableList()
                                novaLista[index] = itemAtualizado
                                mapaProducao = mapaProducao.copy(itens = novaLista)
                            }
                        }
                        CacheManager.salvarProducao(this, mapaProducao)
                    } catch (e: Exception) {
                        e.printStackTrace()
                    }
                },
                onCalculadoraRequest = { produtoNome, tipoOperacao ->
                    try {
                        isTabletMode = isTabletLandscape()
                        if (isTabletMode) {
                            abrirCalculadoraTablet(produtoNome, tipoOperacao, 1)
                        } else {
                            val bottomSheet = CalculadoraBottomSheet.newInstance(produtoNome, tipoOperacao, 1)
                            bottomSheet.setListener(object : CalculadoraBottomSheet.CalculadoraListener {
                                override fun onValorConfirmado(produtoNome: String, tipoOperacao: String, valor: Int, producaoIndex: Int) {
                                    currentAdapter.atualizarQuantidade(produtoNome, valor)
                                }
                            })
                            bottomSheet.show(supportFragmentManager, "calculadora")
                        }
                    } catch (e: Exception) {
                        e.printStackTrace()
                        Toast.makeText(this@MainActivity, "Erro ao abrir calculadora", Toast.LENGTH_SHORT).show()
                    }
                },
                onConfirmarProducao = { item, quantidade ->
                    confirmarProducaoIndividual(item, quantidade)
                }
            )

            recycler.adapter = currentAdapter
        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(this, "Erro ao carregar produtos", Toast.LENGTH_SHORT).show()
        }
    }

    // FUNÇÃO ÚNICA - REMOVER A DUPLICADA QUE ESTÁ MAIS ABAIXO
    private fun confirmarProducaoIndividual(item: ItemProducao, quantidade: Int) {
        try {
            val index = mapaProducao.itens.indexOfFirst {
                it.produto == item.produto && it.categoria == item.categoria
            }

            if (index != -1) {
                val novaLista = mapaProducao.itens.toMutableList()
                val itemExistente = novaLista[index]

                // Usar a nova função para adicionar produção
                val horaAtual = SimpleDateFormat("HH:mm", Locale.getDefault()).format(Date())
                val itemAtualizado = itemExistente.copy(
                    producoes = itemExistente.producoes.toMutableList()
                ).apply {
                    adicionarProducao(quantidade, horaAtual)
                }

                novaLista[index] = itemAtualizado
                mapaProducao = mapaProducao.copy(itens = novaLista)

                // Encontrar o índice da produção (para salvar no Excel)
                val producaoIndex = itemAtualizado.producoes.indexOfFirst {
                    it.quantidade == quantidade && it.hora == horaAtual
                } + 1

                val sucesso = ExcelService.salvarProducao(this, mapaProducao, producaoIndex)

                if (sucesso) {
                    CacheManager.salvarProducao(this, mapaProducao)
                    val dataAtual = SimpleDateFormat("dd/MM/yyyy", Locale.getDefault()).format(Date())
                    findViewById<TextView>(R.id.statusText)?.text = "Produção - $dataAtual"
                    Toast.makeText(this, "Produção de ${item.produto} confirmada!", Toast.LENGTH_SHORT).show()
                    currentAdapter.notifyItemChanged(index)
                } else {
                    Toast.makeText(this, "Erro ao salvar no Excel", Toast.LENGTH_SHORT).show()
                }
            }
        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(this, "Erro ao confirmar produção", Toast.LENGTH_SHORT).show()
        }
    }

    private fun atualizarHeaderUtilizador() {
        val headerView = navigationView.getHeaderView(0)
        val NApp = headerView.findViewById<TextView>(R.id.NApp)
    }

    fun abrirInventario() {
        setContentView(R.layout.activity_inventario)
        findViewById<ImageButton>(R.id.btnVoltar)?.setOnClickListener {
            voltarParaHome()
        }
    }

    private fun abrirResumo() {
        val intent = Intent(this, ResumoActivity::class.java)
        startActivity(intent)
    }

    private fun abrirConfiguracoes() {
        val intent = Intent(this, ConfigActivity::class.java)
        startActivity(intent)
    }

    private fun configurarBotoesHome() {
        findViewById<Button>(R.id.btnProducao).setOnClickListener {
            abrirProducao()
        }

        findViewById<Button>(R.id.btnInventario).setOnClickListener {
            abrirInventario()
        }
    }

    private fun atualizarMensagemBoasVindas() {
        val textViewEscolha = findViewById<TextView>(R.id.textViewEscolhaOpcao)
        textViewEscolha.text = "Bem-vindo ao NataBase"
    }

    fun voltarParaHome() {
        setContentView(R.layout.activity_home)

        drawerLayout = findViewById(R.id.drawerLayout)
        navigationView = findViewById(R.id.navigationView)

        findViewById<ImageButton>(R.id.btnMenu).setOnClickListener {
            drawerLayout.openDrawer(GravityCompat.START)
        }

        atualizarHeaderUtilizador()
        navigationView.setNavigationItemSelectedListener { menuItem ->
            when (menuItem.itemId) {
                R.id.nav_producao -> {
                    abrirProducao()
                    drawerLayout.closeDrawer(GravityCompat.START)
                    true
                }
                R.id.nav_resumo -> {
                    abrirResumo()
                    drawerLayout.closeDrawer(GravityCompat.START)
                    true
                }
                R.id.nav_inventario -> {
                    abrirInventario()
                    drawerLayout.closeDrawer(GravityCompat.START)
                    true
                }
                R.id.nav_configuracoes -> {
                    abrirConfiguracoes()
                    drawerLayout.closeDrawer(GravityCompat.START)
                    true
                }
                else -> false
            }
        }

        atualizarMensagemBoasVindas()
        configurarBotoesHome()
    }

    fun voltarParaCategorias() {
        carregarCategorias()
        findViewById<ImageButton>(R.id.btnVoltar).setOnClickListener {
            voltarParaHome()
        }
    }

    private fun isTabletLandscape(): Boolean {
        val configuration = resources.configuration
        return (configuration.screenLayout and Configuration.SCREENLAYOUT_SIZE_MASK) >= Configuration.SCREENLAYOUT_SIZE_LARGE &&
                configuration.orientation == Configuration.ORIENTATION_LANDSCAPE
    }

    private fun limparProducoes() {
        ExcelService.limparProducoes(this)
        producaoAtualIndex = 1
        CacheManager.salvarProducaoIndex(this, producaoAtualIndex)
        mapaProducao = ExcelService.carregarMapaProducao(this)
        CacheManager.salvarProducao(this, mapaProducao)
        Toast.makeText(this, "Produções limpas com sucesso!", Toast.LENGTH_SHORT).show()
        carregarCategorias()
    }

    private fun regenerarExcelCompleto() {
        mapaProducao = ExcelService.regenerarExcelCompleto(this)
        producaoAtualIndex = 1
        CacheManager.salvarProducaoIndex(this, producaoAtualIndex)
        CacheManager.salvarProducao(this, mapaProducao)
        Toast.makeText(this, "Novo Excel criado com estrutura correta!", Toast.LENGTH_SHORT).show()
        carregarCategorias()
    }

    private fun carregarCategorias() {
        val categorias = mapaProducao.itens
            .map { it.categoria }
            .filter { it.isNotBlank() }
            .distinct()

        val layoutManager = GridLayoutManager(this, calcularNumeroColunas())
        recycler.layoutManager = layoutManager
        recycler.adapter = CategoriaAdapter(categorias) { categoria ->
            abrirProdutosDaCategoria(categoria)
        }

        val dataAtual = SimpleDateFormat("dd/MM/yyyy", Locale.getDefault()).format(Date())
        findViewById<TextView>(R.id.statusText)?.text = "Produção - $dataAtual"
    }

    private fun calcularNumeroColunas(): Int {
        val displayMetrics = resources.displayMetrics
        val screenWidthDp = displayMetrics.widthPixels / displayMetrics.density
        return if (screenWidthDp < 600) 1 else 2
    }

    private fun abrirCalculadoraTablet(produtoNome: String, tipoOperacao: String, producaoIndex: Int) {
        val calculadoraContainer = findViewById<FrameLayout>(R.id.calculadoraContainer)
        val mainContainer = findViewById<LinearLayout>(R.id.mainContainer)

        if (!calculadoraAberta) {
            val calculadoraFragment = CalculadoraFragment.newInstance(produtoNome, tipoOperacao, producaoIndex)
            calculadoraFragment.setListener(this)

            supportFragmentManager.beginTransaction()
                .replace(R.id.calculadoraContainer, calculadoraFragment)
                .commit()

            calculadoraContainer.visibility = View.VISIBLE
            calculadoraContainer.startAnimation(AnimationUtils.loadAnimation(this, R.anim.slide_in_right))

            val mainLayoutParams = mainContainer.layoutParams as LinearLayout.LayoutParams
            mainLayoutParams.weight = 0.7f
            mainContainer.layoutParams = mainLayoutParams

            val calculadoraLayoutParams = calculadoraContainer.layoutParams as LinearLayout.LayoutParams
            calculadoraLayoutParams.weight = 0.3f
            calculadoraContainer.layoutParams = calculadoraLayoutParams

            calculadoraAberta = true
        } else {
            val existingFragment = supportFragmentManager.findFragmentById(R.id.calculadoraContainer) as? CalculadoraFragment
            existingFragment?.let {
                it.updateProductData(produtoNome, tipoOperacao, producaoIndex)
            }
        }
    }

    private fun fecharCalculadoraTablet() {
        if (calculadoraAberta) {
            val calculadoraContainer = findViewById<FrameLayout>(R.id.calculadoraContainer)
            val mainContainer = findViewById<LinearLayout>(R.id.mainContainer)

            calculadoraContainer.startAnimation(AnimationUtils.loadAnimation(this, R.anim.slide_out_right))
            calculadoraContainer.postDelayed({
                val fragment = supportFragmentManager.findFragmentById(R.id.calculadoraContainer)
                fragment?.let {
                    supportFragmentManager.beginTransaction()
                        .remove(it)
                        .commit()
                }

                val mainLayoutParams = mainContainer.layoutParams as LinearLayout.LayoutParams
                mainLayoutParams.weight = 1f
                mainContainer.layoutParams = mainLayoutParams

                val calculadoraLayoutParams = calculadoraContainer.layoutParams as LinearLayout.LayoutParams
                calculadoraLayoutParams.weight = 0f
                calculadoraContainer.layoutParams = calculadoraLayoutParams

                calculadoraContainer.visibility = View.GONE
                calculadoraAberta = false
            }, 300)
        }
    }

    override fun setContentView(layoutResID: Int) {
        isTabletMode = isTabletLandscape() && layoutResID == R.layout.activity_producao

        if (isTabletMode) {
            super.setContentView(R.layout.activity_producao_tablet)
            recycler = findViewById(R.id.recyclerProdutos)

            val drawerLayoutTablet = findViewById<DrawerLayout>(R.id.drawerLayout)
            val navigationViewTablet = findViewById<NavigationView>(R.id.navigationView)

            findViewById<ImageButton>(R.id.btnMenu).setOnClickListener {
                drawerLayoutTablet.openDrawer(GravityCompat.START)
            }

            navigationViewTablet.setNavigationItemSelectedListener { menuItem ->
                when (menuItem.itemId) {
                    R.id.nav_producao -> {
                        drawerLayoutTablet.closeDrawer(GravityCompat.START)
                        true
                    }
                    R.id.nav_inventario -> {
                        abrirInventario()
                        drawerLayoutTablet.closeDrawer(GravityCompat.START)
                        true
                    }
                    R.id.nav_configuracoes -> {
                        abrirConfiguracoes()
                        drawerLayoutTablet.closeDrawer(GravityCompat.START)
                        true
                    }
                    else -> false
                }
            }

            val calculadoraContainer = findViewById<FrameLayout>(R.id.calculadoraContainer)
            calculadoraContainer.visibility = View.GONE
            calculadoraAberta = false
        } else {
            super.setContentView(layoutResID)
            if (layoutResID == R.layout.activity_producao) {
                recycler = findViewById(R.id.recyclerProdutos)
                adicionarBotaoDebug()
            }
        }
    }

    override fun onValorConfirmado(produtoNome: String, tipoOperacao: String, valor: Int, producaoIndex: Int) {
        currentAdapter.atualizarQuantidade(produtoNome, valor)
        fecharCalculadoraTablet()
    }

    override fun onCalculadoraFechada() {
        fecharCalculadoraTablet()
    }

    private fun calcularNumeroColunasProdutos(): Int {
        val displayMetrics = resources.displayMetrics
        val screenWidthDp = displayMetrics.widthPixels / displayMetrics.density

        return when {
            screenWidthDp < 600 -> 1
            screenWidthDp < 900 -> 2
            else -> 3
        }
    }

    private fun adicionarBotaoDebug() {
        var topBar: LinearLayout? = findViewById(R.id.topBar)

        if (topBar == null) {
            val mainContainer = findViewById<LinearLayout?>(R.id.mainContainer)
            topBar = mainContainer?.findViewById(R.id.topBar)
        }

        topBar?.let { bar ->
            val debugLayouts = bar.findViewById<LinearLayout?>(R.id.debugLayout)
            debugLayouts?.let {
                bar.removeView(it)
            }

            val debugLayout = LinearLayout(this).apply {
                id = R.id.debugLayout
                orientation = LinearLayout.HORIZONTAL
                layoutParams = LinearLayout.LayoutParams(
                    LinearLayout.LayoutParams.WRAP_CONTENT,
                    LinearLayout.LayoutParams.WRAP_CONTENT
                ).apply {
                    gravity = android.view.Gravity.END
                }
            }

            val limparButton = Button(this).apply {
                text = "Limpar"
                setBackgroundResource(R.drawable.button_back)
                setTextColor(ContextCompat.getColor(this@MainActivity, R.color.white))
                layoutParams = LinearLayout.LayoutParams(
                    LinearLayout.LayoutParams.WRAP_CONTENT,
                    LinearLayout.LayoutParams.WRAP_CONTENT
                ).apply {
                    setMargins(0, 0, 8, 0)
                }
                setOnClickListener {
                    mostrarDialogoLimpeza()
                }
            }

            val regenerarButton = Button(this).apply {
                text = "Novo Excel"
                setBackgroundResource(R.drawable.button_back)
                setTextColor(ContextCompat.getColor(this@MainActivity, R.color.white))
                layoutParams = LinearLayout.LayoutParams(
                    LinearLayout.LayoutParams.WRAP_CONTENT,
                    LinearLayout.LayoutParams.WRAP_CONTENT
                )
                setOnClickListener {
                    mostrarDialogoRegeneracao()
                }
            }

            debugLayout.addView(limparButton)
            debugLayout.addView(regenerarButton)
            bar.addView(debugLayout)
        }
    }

    private fun mostrarDialogoLimpeza() {
        AlertDialog.Builder(this)
            .setTitle("Limpar Produções")
            .setMessage("Tem certeza que deseja limpar TODAS as produções? Esta ação não pode ser desfeita.")
            .setPositiveButton("Sim") { _, _ ->
                limparProducoes()
            }
            .setNegativeButton("Cancelar", null)
            .show()
    }

    private fun mostrarDialogoRegeneracao() {
        AlertDialog.Builder(this)
            .setTitle("Criar Novo Excel")
            .setMessage("Tem certeza que deseja criar um NOVO arquivo Excel? Isto apagará o arquivo atual e criará um novo com a estrutura correta.")
            .setPositiveButton("Sim") { _, _ ->
                regenerarExcelCompleto()
            }
            .setNegativeButton("Cancelar", null)
            .show()
    }

    override fun onResume() {
        super.onResume()
        producaoAtualIndex = CacheManager.carregarProducaoIndex(this)
    }

    override fun onConfigurationChanged(newConfig: Configuration) {
        super.onConfigurationChanged(newConfig)

        if (::recycler.isInitialized) {
            val currentAdapter = recycler.adapter
            if (currentAdapter is CategoriaAdapter) {
                carregarCategorias()
            } else if (currentAdapter is ItemProducaoAdapter) {
                abrirProdutosDaCategoria(categoriaAtual)
            }
        }
    }
}