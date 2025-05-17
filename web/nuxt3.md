# Nuxt 3 - The Progressive Web Framework

## Che cos'è Nuxt 3
Nuxt 3 è un framework web full-stack basato su Vue.js che semplifica lo sviluppo di applicazioni web moderne. È progettato per essere progressivo, permettendo di attivare funzionalità man mano che il progetto cresce.

## Caratteristiche Principali

### 1. Architettura Progressiva
- **Server-side rendering (SSR)**: Rendering server-side per prestazioni ottimali
- **Static Site Generation (SSG)**: Generazione statica per siti statici
- **Hybrid Rendering**: Combinazione di SSR e SSG
- **Client-side rendering**: Rendering lato client

### 2. Funzionalità di Base
```javascript
// Struttura di base di un progetto Nuxt 3
├── app.vue          # Layout principale
├── pages/           # Pagine dell'applicazione
├── components/      # Componenti Vue
├── composables/     # Logica condivisa
├── middleware/      # Middleware HTTP
├── server/          # API e middleware server
├── public/          # File statici
├── plugins/         # Plugins personalizzati
├── stores/          # Pinia Stores
└── nuxt.config.js   # Configurazione
```

### 3. Routing
- **File-based routing**: Le pagine sono create automaticamente dai file nella cartella `pages`
- **Dynamic routes**: Supporto per route dinamiche
- **Nested routes**: Route annidate
- **Middleware**: Gestione delle route

### 4. Composizione e Gestione Stato

#### Composables
```javascript
// composables/useApi.js
export function useApi() {
  const { $axios } = useNuxtApp()

  const fetchData = async (endpoint) => {
    try {
      const response = await $axios.get(endpoint)
      return response.data
    } catch (error) {
      console.error('Errore durante il fetch:', error)
      throw error
    }
  }

  return {
    fetchData
  }
}

// composables/useAuth.js
export function useAuth() {
  const user = useState('user', () => null)
  const token = useState('token', () => null)

  const login = async (credentials) => {
    const { data } = await useFetch('/api/auth/login', {
      method: 'POST',
      body: credentials
    })
    
    user.value = data.value.user
    token.value = data.value.token
  }

  const logout = () => {
    user.value = null
    token.value = null
  }

  return {
    user,
    token,
    login,
    logout
  }
}
```

#### Pinia Stores
```javascript
// stores/user.js
import { defineStore } from 'pinia'

export const useUserStore = defineStore('user', {
  state: () => ({
    user: null,
    token: null
  }),
  actions: {
    setUser(user) {
      this.user = user
    },
    setToken(token) {
      this.token = token
    },
    clear() {
      this.user = null
      this.token = null
    }
  }
})

// stores/products.js
import { defineStore } from 'pinia'
import { useApi } from '~/composables/useApi'

export const useProductsStore = defineStore('products', {
  state: () => ({
    products: [],
    loading: false,
    error: null
  }),
  actions: {
    async fetchProducts() {
      this.loading = true
      this.error = null
      
      try {
        const api = useApi()
        this.products = await api.fetchData('/api/products')
      } catch (err) {
        this.error = err.message
      } finally {
        this.loading = false
      }
    }
  }
})
```

#### Utilizzo nei Componenti
```vue
<!-- pages/index.vue -->
<template>
  <v-container>
    <v-row>
      <v-col>
        <v-card>
          <v-card-title>Benvenuto</v-card-title>
          <v-card-text>
            <v-btn color="primary" @click="fetchProducts">
              Carica Prodotti
            </v-btn>
          </v-card-text>
        </v-card>
      </v-col>
    </v-row>
  </v-container>
</template>

<script>
import { useProductsStore } from '~/stores/products'
import { useUserStore } from '~/stores/user'

export default {
  setup() {
    const productsStore = useProductsStore()
    const userStore = useUserStore()

    const fetchProducts = () => {
      productsStore.fetchProducts()
    }

    return {
      fetchProducts
    }
  }
}
</script>

<style scoped>
/* Stili locali */
</style>
```

### 5. Server-side
- **API Routes**: Gestione delle API
- **Middleware**: Middleware server
- **Server Components**: Componenti server-side
- **Server Functions**: Funzioni server-side

### 6. Plugins e ACL

#### Plugins
```javascript
// plugins/axios.js
export default defineNuxtPlugin((nuxtApp) => {
  nuxtApp.vueApp.config.globalProperties.$axios = nuxtApp.$axios
})

// plugins/acl.js
import { defineNuxtPlugin } from '#app'
import { useAuthStore } from '~/stores/auth'

export default defineNuxtPlugin((nuxtApp) => {
  const authStore = useAuthStore()
  
  nuxtApp.vueApp.config.globalProperties.$hasPermission = (permission) => {
    if (!authStore.user) return false
    return authStore.user.permissions.includes(permission)
  }
})
```

#### ACL Store
```javascript
// stores/auth.js
import { defineStore } from 'pinia'

export const useAuthStore = defineStore('auth', {
  state: () => ({
    user: null,
    token: null,
    permissions: []
  }),
  actions: {
    setPermissions(permissions) {
      this.permissions = permissions
    },
    hasPermission(permission) {
      return this.permissions.includes(permission)
    }
  }
})
```

#### Utilizzo in Componenti
```vue
<!-- pages/dashboard.vue -->
<template>
  <v-container>
    <v-row>
      <v-col>
        <v-card v-if="$hasPermission('admin:dashboard')">
          <v-card-title>Pannello Amministratore</v-card-title>
          <v-card-text>
            <v-btn color="primary" @click="fetchStats">
              Carica Statistiche
            </v-btn>
          </v-card-text>
        </v-card>
      </v-col>
    </v-row>
  </v-container>
</template>

<script>
import { useAuthStore } from '~/stores/auth'

export default {
  setup() {
    const authStore = useAuthStore()
    
    const fetchStats = async () => {
      try {
        const { data } = await useFetch('/api/stats', {
          headers: {
            Authorization: `Bearer ${authStore.token}`
          }
        })
        // Gestione dati
      } catch (error) {
        console.error('Errore:', error)
      }
    }

    return {
      fetchStats
    }
  }
}
</script>
```

### 6. Performance
- **Automatic Code Splitting**: Splitting del codice automatico
- **Tree Shaking**: Rimozione del codice non utilizzato
- **Lazy Loading**: Caricamento lazy dei componenti
- **Static Asset Optimization**: Ottimizzazione degli asset statici

### 7. Configurazione
```javascript
// nuxt.config.js
export default defineNuxtConfig({
  // UI Framework e Moduli
  modules: [
    '@nuxtjs/vuetify',
    '@pinia/nuxt',
    '@nuxtjs/axios',
    '@nuxtjs/auth-next'
  ],

  // Vuetify Configuration
  vuetify: {
    theme: {
      defaultTheme: 'light',
      themes: {
        light: {
          colors: {
            primary: '#1867c0',
            secondary: '#424242',
            accent: '#82B1FF',
            error: '#FF5252',
            info: '#2196F3',
            success: '#4CAF50',
            warning: '#FFC107'
          }
        }
      }
    }
  },

  // Configurazione Axios
  axios: {
    baseURL: process.env.NUXT_PUBLIC_API_BASE || 'http://localhost:3000/api'
  },

  // Configurazione Autenticazione
  auth: {
    strategies: {
      local: {
        token: {
          property: 'token',
          global: true,
          maxAge: 86400
        },
        user: {
          property: 'user',
          autoFetch: true
        },
        endpoints: {
          login: { url: '/auth/login', method: 'post' },
          logout: { url: '/auth/logout', method: 'post' },
          user: { url: '/auth/user', method: 'get' }
        }
      }
    }
  },

  // Plugins
  plugins: [
    '~/plugins/acl.js',
    '~/plugins/axios.js',
    '~/plugins/vuetify.js'
  ],

  // App Configuration
  app: {
    head: {
      title: 'Mio App',
      meta: [
        { name: 'description', content: 'Descrizione dell'app' }
      ]
    }
  },

  // Runtime Configuration
  runtimeConfig: {
    public: {
      apiBase: process.env.NUXT_PUBLIC_API_BASE
    },
    private: {
      jwtSecret: process.env.JWT_SECRET
    }
  }
})
```

## Best Practices

1. **Organizzazione del Codice**
   - Mantieni componenti piccoli e riutilizzabili
   - Usa composables per la logica condivisa
   - Organizza le route con file-based routing

2. **Gestione degli Stati**
   - Usa Pinia per la gestione dello stato
   - Implementa middleware per la validazione
   - Usa composables per la logica condivisa

3. **Performance**
   - Implementa lazy loading per i componenti
   - Ottimizza le immagini e gli asset
   - Usa server-side rendering quando possibile

4. **SEO**
   - Usa meta tags appropriati
   - Implementa server-side rendering
   - Gestisci i redirect e gli errori 404

## Esempio di Applicazione
```vue
<!-- app.vue -->
<template>
  <div>
    <NuxtLayout>
      <NuxtPage />
    </NuxtLayout>
  </div>
</template>

<!-- pages/index.vue -->
<template>
  <div>
    <h1>Benvenuto</h1>
    <NuxtLink to="/about">About</NuxtLink>
  </div>
</template>

<!-- composables/useApi.ts -->
export const useApi = () => {
  const fetchApi = async (endpoint) => {
    const { data } = await useFetch(endpoint)
    return data.value
  }
  return { fetchApi }
}
```

## Moduli Utili

### 1. UI Framework
- `@nuxtjs/tailwindcss`: Tailwind CSS
- `@nuxtjs/vuetify`: Vuetify 3
- `@nuxtjs/primevue`: PrimeVue

### 2. Gestione Stato
- `@pinia/nuxt`: Pinia
- `@nuxtjs/composition-api`: Composition API

### 3. API e Dati
- `@nuxtjs/axios`: Axios
- `@nuxtjs/supabase`: Supabase
- `@nuxt/content`: Gestione contenuti

### 4. Autenticazione
- `@nuxtjs/auth-next`: Autenticazione
- `@nuxtjs/firebase`: Firebase
- `@nuxtjs/supabase`: Supabase

## Deployment

### 1. Hosting
- **Vercel**: Hosting nativo
- **Netlify**: Deploy automatico
- **AWS**: Deploy su AWS
- **Docker**: Containerizzazione

### 2. Configurazione
```typescript
// nuxt.config.ts
export default defineNuxtConfig({
  // Configurazione per il deployment
  nitro: {
    preset: 'vercel',
    compressPublicAssets: true
  },

  // Environment variables
  runtimeConfig: {
    public: {
      apiBase: process.env.NUXT_PUBLIC_API_BASE
    }
  }
})
```

## Miglioramenti rispetto a Nuxt 2

1. **Performance**
   - Build più veloce
   - Ottimizzazione automatica
   - Miglior supporto per il lazy loading

2. **API**
   - API più moderna
   - Supporto per TypeScript
   - Miglior gestione degli errori

3. **Developer Experience**
   - Hot Module Replacement migliorato
   - Miglior feedback degli errori
   - Configurazione più intuitiva

4. **Modularità**
   - Supporto per moduli più potente
   - Isolamento dei moduli
   - Gestione delle dipendenze migliorata

## Consigli per lo Sviluppo

1. **Testing**
   - Usa Vitest per i test unitari
   - Implementa test end-to-end
   - Usa composables per i test

2. **Debug**
   - Abilita il modo sviluppo
   - Usa la console di Vue DevTools
   - Implementa logging

3. **Sicurezza**
   - Validazione dei dati
   - Sanitizzazione del contenuto
   - Gestione delle sessioni

4. **Performance**
   - Implementa caching
   - Ottimizza le immagini
   - Usa lazy loading

## Conclusione
Nuxt 3 rappresenta un grande passo avanti nel mondo del development web, offrendo un'esperienza di sviluppo moderna e produttiva basata su Vue.js. La sua natura progressiva e modulare lo rende adatto a progetti di qualsiasi dimensione, dalla semplice landing page alle complesse applicazioni enterprise.
