package com.belasoft.natabase_alpha

import android.os.Bundle
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.Button
import android.widget.TextView
import androidx.fragment.app.Fragment

class CalculadoraFragment : Fragment() {

    private lateinit var tvDisplay: TextView
    private lateinit var tvProdutoSelecionado: TextView
    private var currentValue = 0
    private var produtoNome: String = ""
    private var tipoOperacao: String = ""
    private var producaoIndex: Int = 0

    interface CalculadoraListener {
        fun onValorConfirmado(produtoNome: String, tipoOperacao: String, valor: Int, producaoIndex: Int)
        fun onCalculadoraFechada()
    }

    private var listener: CalculadoraListener? = null

    companion object {
        fun newInstance(produtoNome: String, tipoOperacao: String, producaoIndex: Int): CalculadoraFragment {
            val args = Bundle().apply {
                putString("PRODUTO_NOME", produtoNome)
                putString("TIPO_OPERACAO", tipoOperacao)
                putInt("PRODUCAO_INDEX", producaoIndex)
            }
            return CalculadoraFragment().apply {
                arguments = args
            }
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        arguments?.let {
            produtoNome = it.getString("PRODUTO_NOME") ?: ""
            tipoOperacao = it.getString("TIPO_OPERACAO") ?: ""
            producaoIndex = it.getInt("PRODUCAO_INDEX", 0)
        }
    }

    override fun onCreateView(
        inflater: LayoutInflater,
        container: ViewGroup?,
        savedInstanceState: Bundle?
    ): View? {
        return inflater.inflate(R.layout.fragment_calculadora, container, false)
    }

    override fun onViewCreated(view: View, savedInstanceState: Bundle?) {
        super.onViewCreated(view, savedInstanceState)

        tvDisplay = view.findViewById(R.id.tvDisplay)
        tvProdutoSelecionado = view.findViewById(R.id.tvProdutoSelecionado)

        updateProdutoInfo()
        updateDisplay()
        setupClickListeners(view)
    }

    fun updateProdutoInfo() {
        val textoOperacao = when (tipoOperacao) {
            "Produção" -> "Produção ${if (producaoIndex > 0) producaoIndex else ""}"
            "Perdas" -> "Perdas"
            "Sobras" -> "Sobras"
            else -> tipoOperacao
        }
        tvProdutoSelecionado.text = "Produto: $produtoNome - $textoOperacao"
    }

    fun updateProductData(newProdutoNome: String, newTipoOperacao: String, newProducaoIndex: Int) {
        produtoNome = newProdutoNome
        tipoOperacao = newTipoOperacao
        producaoIndex = newProducaoIndex
        updateProdutoInfo()
        clear()
    }

    private fun setupClickListeners(view: View) {
        view.findViewById<Button>(R.id.btn0).setOnClickListener { appendNumber(0) }
        view.findViewById<Button>(R.id.btn1).setOnClickListener { appendNumber(1) }
        view.findViewById<Button>(R.id.btn2).setOnClickListener { appendNumber(2) }
        view.findViewById<Button>(R.id.btn3).setOnClickListener { appendNumber(3) }
        view.findViewById<Button>(R.id.btn4).setOnClickListener { appendNumber(4) }
        view.findViewById<Button>(R.id.btn5).setOnClickListener { appendNumber(5) }
        view.findViewById<Button>(R.id.btn6).setOnClickListener { appendNumber(6) }
        view.findViewById<Button>(R.id.btn7).setOnClickListener { appendNumber(7) }
        view.findViewById<Button>(R.id.btn8).setOnClickListener { appendNumber(8) }
        view.findViewById<Button>(R.id.btn9).setOnClickListener { appendNumber(9) }

        view.findViewById<Button>(R.id.btnClear).setOnClickListener { clear() }
        view.findViewById<Button>(R.id.btnBackspace).setOnClickListener { backspace() }

        view.findViewById<Button>(R.id.btnConfirmar).setOnClickListener {
            listener?.onValorConfirmado(produtoNome, tipoOperacao, currentValue, producaoIndex)
            listener?.onCalculadoraFechada()
        }

        view.findViewById<Button>(R.id.btnFechar).setOnClickListener {
            listener?.onCalculadoraFechada()
        }
    }

    private fun appendNumber(number: Int) {
        currentValue = currentValue * 10 + number
        updateDisplay()
    }

    fun clear() {
        currentValue = 0
        updateDisplay()
    }

    private fun backspace() {
        currentValue = currentValue / 10
        updateDisplay()
    }

    private fun updateDisplay() {
        tvDisplay.text = currentValue.toString()
    }

    fun setListener(listener: CalculadoraListener) {
        this.listener = listener
    }
}