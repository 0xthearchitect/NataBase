package com.belasoft.natabase_alpha

import android.content.Intent
import android.os.Bundle
import android.view.View
import android.widget.Button
import android.widget.ImageButton
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.GravityCompat
import androidx.drawerlayout.widget.DrawerLayout
import com.google.android.material.navigation.NavigationView

abstract class BaseActivity : AppCompatActivity() {

    protected lateinit var drawerLayout: DrawerLayout
    protected lateinit var navigationView: NavigationView

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        // As classes filhas chamarão setContentView com seu próprio layout
    }

    protected fun setupDrawer() {
        try {
            drawerLayout = findViewById(R.id.drawerLayout)
            navigationView = findViewById(R.id.navigationView)

            // Configurar botão do menu
            val btnMenu = findViewById<ImageButton>(R.id.btnMenu)
            btnMenu?.setOnClickListener {
                drawerLayout.openDrawer(GravityCompat.START)
            }

            // Configurar navegação
            navigationView.setNavigationItemSelectedListener { menuItem ->
                handleNavigation(menuItem.itemId)
                drawerLayout.closeDrawer(GravityCompat.START)
                true
            }

            // Atualizar header
            updateUserHeader()

            // Marcar item atual como selecionado
            markCurrentItem()

        } catch (e: Exception) {
            // Log do erro sem crashar a app
            e.printStackTrace()
        }
    }

    private fun handleNavigation(menuItemId: Int) {
        when (menuItemId) {
            R.id.nav_home -> {
                if (this !is MainActivity) {
                    val intent = Intent(this, MainActivity::class.java)
                    intent.flags = Intent.FLAG_ACTIVITY_CLEAR_TOP or Intent.FLAG_ACTIVITY_SINGLE_TOP
                    startActivity(intent)
                    if (this !is MainActivity) finish()
                } else {
                    // Se já estamos na MainActivity, garantir que está na home
                    (this as MainActivity).voltarParaHome()
                }
            }

            R.id.nav_producao -> {
                if (this is MainActivity) {
                    // Se já estamos na MainActivity, apenas muda para produção
                    (this as MainActivity).abrirProducao()
                } else {
                    // Se estamos em outra activity, vai para MainActivity e abre produção
                    val intent = Intent(this, MainActivity::class.java)
                    intent.flags = Intent.FLAG_ACTIVITY_CLEAR_TOP or Intent.FLAG_ACTIVITY_SINGLE_TOP
                    intent.putExtra("OPEN_PRODUCTION", true)
                    startActivity(intent)
                    finish()
                }
            }

            R.id.nav_resumo -> {
                if (this !is ResumoActivity) {
                    val intent = Intent(this, ResumoActivity::class.java)
                    startActivity(intent)
                }
            }

            R.id.nav_inventario -> {
                if (this is MainActivity) {
                    // Se já estamos na MainActivity, apenas muda para inventário
                    (this as MainActivity).abrirInventario()
                } else {
                    // Se estamos em outra activity, vai para MainActivity e abre inventário
                    val intent = Intent(this, MainActivity::class.java)
                    intent.flags = Intent.FLAG_ACTIVITY_CLEAR_TOP or Intent.FLAG_ACTIVITY_SINGLE_TOP
                    intent.putExtra("OPEN_INVENTORY", true)
                    startActivity(intent)
                    finish()
                }
            }

            R.id.nav_configuracoes -> {
                if (this !is ConfigActivity) {
                    val intent = Intent(this, ConfigActivity::class.java)
                    startActivity(intent)
                }
            }
        }
    }

    private fun updateUserHeader() {
        try {
            val headerView = navigationView.getHeaderView(0)
            val NApp = headerView.findViewById<TextView>(R.id.NApp)

            NApp.text = "NataBase"
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }

    private fun markCurrentItem() {
        try {
            when (this) {
                is MainActivity -> {
                    // Verifica se estamos na home ou na produção
                    val isHome = findViewById<Button>(R.id.btnProducao) != null
                    if (isHome) {
                        navigationView.setCheckedItem(R.id.nav_home)
                    } else {
                        navigationView.setCheckedItem(R.id.nav_producao)
                    }
                }
                is ResumoActivity -> navigationView.setCheckedItem(R.id.nav_resumo)
                is ConfigActivity -> navigationView.setCheckedItem(R.id.nav_configuracoes)
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }

    // Método para esconder/mostrar o drawer quando necessário
    fun setDrawerEnabled(enabled: Boolean) {
        try {
            if (enabled) {
                drawerLayout.setDrawerLockMode(DrawerLayout.LOCK_MODE_UNLOCKED)
                findViewById<ImageButton>(R.id.btnMenu)?.visibility = View.VISIBLE
            } else {
                drawerLayout.setDrawerLockMode(DrawerLayout.LOCK_MODE_LOCKED_CLOSED)
                findViewById<ImageButton>(R.id.btnMenu)?.visibility = View.GONE
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }
}