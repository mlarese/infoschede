# Componenti Vuetify 3

## Componenti Principali

### 1. Layout Components

#### v-container
**Descrizione**: Contenitore principale che gestisce il layout della pagina
**Funzione**: Crea un contenitore responsivo con margini standard
```vue
<!-- Attributi principali -->
<v-container
  fluid="true"        <!-- Layout fluido -->
  class="my-class"    <!-- Classe personalizzata -->
  style="margin: 20px" <!-- Stili personalizzati -->
  max-width="1200"    <!-- Larghezza massima -->
>
  <v-row>
    <v-col
      cols="12"       <!-- Colonne su mobile -->
      sm="6"          <!-- Colonne su tablet -->
      md="4"          <!-- Colonne su desktop -->
      lg="3"          <!-- Colonne su grandi schermi -->
      xl="2"          <!-- Colonne su schermi extra-large -->
      
      /* Spaziatura */
      class="pa-4"    <!-- Padding -->
      :offset="0"     <!-- Spostamento da sinistra -->
      :offset-sm="0"  <!-- Spostamento su tablet -->
      :offset-md="0"  <!-- Spostamento su desktop -->
      :offset-lg="0"  <!-- Spostamento su grandi schermi -->
      :offset-xl="0"  <!-- Spostamento su schermi extra-large -->
      
      /* Allineamento */
      align="start"   <!-- Allineamento verticale -->
      align-self="start" <!-- Allineamento verticale specifico -->
      
      /* Ordine */
      :order="0"      <!-- Ordine di visualizzazione -->
      :order-sm="0"   <!-- Ordine su tablet -->
      :order-md="0"   <!-- Ordine su desktop -->
      :order-lg="0"   <!-- Ordine su grandi schermi -->
      :order-xl="0"   <!-- Ordine su schermi extra-large -->
      
      /* Spaziatura interna */
      :padding-x="0"  <!-- Padding orizzontale -->
      :padding-y="0"  <!-- Padding verticale -->
      :margin-x="0"   <!-- Margin orizzontale -->
      :margin-y="0"   <!-- Margin verticale -->
      
      /* Responsive */
      :cols="12"      <!-- Colonne -->
      :sm="6"        <!-- Colonne su tablet -->
      :md="4"        <!-- Colonne su desktop -->
      :lg="3"        <!-- Colonne su grandi schermi -->
      :xl="2"        <!-- Colonne su schermi extra-large -->
      
      /* Stato */
      :disabled="false" <!-- Disabilitato -->
      :hidden="false"   <!-- Nascosto -->
      
      /* Layout */
      :no-gutters="false" <!-- Senza spaziature -->
      :reverse="false"    <!-- Inverso -->
    >
      <!-- Contenuto -->
    </v-col>
  </v-row>
</v-container>
```

#### v-navigation-drawer
**Descrizione**: Menu laterale navigabile
**Funzione**: Fornisce un'area di navigazione che può essere mostrata/ascosta
```vue
<!-- Attributi principali -->
<v-navigation-drawer
  v-model="drawer"     <!-- Stato aperto/chiuso -->
  permanent="false"    <!-- Tipo di drawer -->
  temporary="false"    <!-- Drawer temporaneo -->
  clipped="false"      <!-- Clipping -->
  floating="false"     <!-- Floating -->
  mini-variant="false" <!-- Mini variant -->
  expand-on-hover="false" <!-- Espansione su hover -->
  app="false"         <!-- App drawer -->
  disable-route-watcher="false" <!-- Disabilita il watcher -->
>
  <v-list density="compact">
    <v-list-item
      prepend-icon="mdi-home"
      title="Home"
      to="/"
      value="home"
      active-color="primary"
    />
  </v-list>
</v-navigation-drawer>
```

#### v-app-bar
**Descrizione**: Barra di navigazione superiore
**Funzione**: Fornisce una barra fissa in alto con titolo e azioni
```vue
<v-app-bar>
  <v-app-bar-title>Titolo App</v-app-bar-title>
</v-app-bar>
```
```

### 2. Form Components

#### v-text-field
**Descrizione**: Campo di input testuale
**Funzione**: Permette l'inserimento di testo con label e validazione
```vue
<!-- Attributi principali -->
<v-text-field
  v-model="nome"
  label="Nome"
  variant="outlined"      <!-- Variant: outlined, filled, solo, underlined -->
  density="default"       <!-- Density: default, compact, comfortable -->
  prepend-icon="mdi-account" <!-- Icona precedente -->
  append-icon="mdi-close" <!-- Icona successiva -->
  clearable="true"        <!-- Bottone di pulizia -->
  readonly="false"       <!-- Campo di sola lettura -->
  disabled="false"       <!-- Campo disabilitato -->
  required="false"       <!-- Campo obbligatorio -->
  :rules="[rules.required]" <!-- Regole di validazione -->
  :error-messages="errors" <!-- Messaggi di errore -->
  :counter="20"         <!-- Contatore caratteri -->
  :loading="false"      <!-- Stato di caricamento -->
  :persistent-hint="true" <!-- Hint persistente -->
  hint="Inserisci il tuo nome" <!-- Hint -->
  :maxlength="50"      <!-- Lunghezza massima -->
  :minlength="2"       <!-- Lunghezza minima -->
  :autofocus="false"   <!-- Autofocus -->
  :persistent-placeholder="false" <!-- Placeholder persistente -->
/>
```

#### v-select
**Descrizione**: Menu a tendina per selezioni
**Funzione**: Permette di scegliere un'opzione da una lista
```vue
<!-- Select Standard -->
<v-select
  v-model="selezione"
  :items="[
    { value: 'op1', title: 'Opzione 1' },
    { value: 'op2', title: 'Opzione 2' }
  ]"
  label="Seleziona"
  variant="outlined"
  density="default"
  :clearable="true"
/>

<!-- Select con selezione multipla -->
<v-select
  v-model="selezioni"
  :items="[
    { value: 'op1', title: 'Opzione 1' },
    { value: 'op2', title: 'Opzione 2' }
  ]"
  label="Selezioni multiple"
  variant="outlined"
  density="default"
  multiple
  chips
  deletable-chips
/>

<!-- Select con ricerca -->
<v-select
  v-model="ricerca"
  :items="items"
  label="Ricerca"
  variant="outlined"
  density="default"
  :search="true"
  :menu-props="{ maxHeight: '400' }"
/>

<!-- Select con caricamento lazy -->
<v-select
  v-model="lazy"
  :items="[]"
  label="Caricamento lazy"
  variant="outlined"
  density="default"
  :eager="false"
  :loading="isLoading"
  @update:search="fetchItems"
/>
```

#### v-textarea
**Descrizione**: Campo di input multiriga
**Funione**: Permette l'inserimento di testo su più righe
```vue
<v-textarea
  v-model="descrizione"
  label="Descrizione"
  variant="outlined"
  density="default"
  :auto-grow="true"
  :rows="3"
  :max-rows="6"
  :counter="200"
  :no-resize="false"
/>
```

#### v-checkbox
**Descrizione**: Casella di spunta
**Funzione**: Permette di selezionare/deselezionare opzioni booleane
```vue
<v-checkbox
  v-model="accetto"
  label="Accetto i termini"
  :indeterminate="false"
  :disabled="false"
  :readonly="false"
  :color="primary"
  :true-value="true"
  :false-value="false"
/>
```

#### v-radio-group
**Descrizione**: Gruppo di opzioni radio
**Funzione**: Permette di selezionare una singola opzione tra più scelte
```vue
<v-radio-group v-model="opzione">
  <v-radio
    v-for="(item, index) in opzioni"
    :key="index"
    :label="item.label"
    :value="item.value"
    :color="primary"
  />
</v-radio-group>
```

#### v-switch
**Descrizione**: Interruttore on/off
**Funzione**: Permette di attivare/disattivare una funzionalità
```vue
<v-switch
  v-model="attivo"
  label="Attivo"
  :color="primary"
  :disabled="false"
  :readonly="false"
  :inset="false"
/>
```

#### v-slider
**Descrizione**: Slider per selezione numerica
**Funzione**: Permette di selezionare un valore numerico su una scala
```vue
<v-slider
  v-model="valore"
  :min="0"
  :max="100"
  :step="1"
  :label="Valore"
  :thumb-label="always"
  :disabled="false"
  :readonly="false"
/>
```

#### v-file-input
**Descrizione**: Campo per caricamento file
**Funzione**: Permette di caricare file
```vue
<v-file-input
  v-model="file"
  label="Carica file"
  :accept=".pdf,.doc,.docx"
  :counter="true"
  :multiple="false"
  :show-size="true"
  :disabled="false"
  :readonly="false"
  :truncated-length="20"
/>
```

#### v-combobox
**Descrizione**: Campo di input con suggerimenti
**Funzione**: Combina input e select con suggerimenti
```vue
<v-combobox
  v-model="valore"
  :items="suggerimenti"
  label="Suggerimenti"
  :search-input.sync="search"
  :hide-selected="true"
  :multiple="false"
  :clearable="true"
/>
```

#### v-autocomplete
**Descrizione**: Campo di input con autocompletamento
**Funzione**: Suggerisce opzioni man mano che si digita
```vue
<v-autocomplete
  v-model="valore"
  :items="opzioni"
  label="Autocompletamento"
  :search-input.sync="search"
  :hide-selected="true"
  :multiple="false"
  :clearable="true"
  :loading="false"
  :no-data-text="Nessun risultato trovato"
/>
```

#### v-checkbox
**Descrizione**: Casella di spunta
**Funzione**: Permette di selezionare/deselezionare opzioni booleane
```vue
<v-checkbox
  v-model="accetto"
  label="Accetto i termini"
/>
```

### 3. Navigation Components

#### v-bottom-navigation
**Descrizione**: Barra di navigazione inferiore
**Funzione**: Fornisce un'interfaccia di navigazione per dispositivi mobili
```vue
<v-bottom-navigation>
  <v-btn value="recent">
    <v-icon>mdi-history</v-icon>
    <span>Recent</span>
  </v-btn>
</v-bottom-navigation>
```

#### v-breadcrumbs
**Descrizione**: Componente di navigazione a briciole di pane
**Funzione**: Mostra la posizione attuale all'interno della gerarchia
```vue
<v-breadcrumbs :items="[
  { title: 'Home', to: '/' },
  { title: 'Dashboard', to: '/dashboard' }
]">
  <template v-slot:divider>
    <v-icon>mdi-chevron-right</v-icon>
  </template>
</v-breadcrumbs>
```
```

### 4. Data Components

#### v-table
**Descrizione**: Tabella per visualizzare dati tabulari
**Funzione**: Mostra dati in formato tabellare con intestazioni
```vue
<!-- Attributi principali -->
<v-table
  density="default"      <!-- Density: default, compact, comfortable -->
  fixed-header="false"   <!-- Header fisso -->
  hover="false"         <!-- Hover sugli elementi -->
  striped="false"       <!-- Strisce alternate -->
  height="auto"         <!-- Altezza -->
  :items-per-page="10"  <!-- Elementi per pagina -->
>
  <thead>
    <tr>
      <th class="text-left">
        Nome
      </th>
      <th class="text-right">
        Valore
      </th>
    </tr>
  </thead>
  <tbody>
    <tr v-for="item in items" :key="item.id">
      <td>{{ item.nome }}</td>
      <td>{{ item.valore }}</td>
    </tr>
  </tbody>
</v-table>
```

#### v-data-iterator
**Descrizione**: Componente per iterare su dati
**Funzione**: Fornisce una struttura per visualizzare e paginare dati
```vue
<v-data-iterator :items="items">
  <template v-slot:default="{ items }">
    <v-item-group>
      <v-item v-for="item in items" :key="item.id">
        {{ item.name }}
      </v-item>
    </v-item-group>
  </template>
</v-data-iterator>
```
```

### 5. Content Components

#### v-card
**Descrizione**: Card per contenuto raggruppato
**Funzione**: Raggruppa contenuti correlati in una card con titolo, testo e azioni
```vue
<v-card>
  <v-card-title>Titolo</v-card-title>
  <v-card-text>Contenuto</v-card-text>
  <v-card-actions>
    <v-btn color="primary">Azione</v-btn>
  </v-card-actions>
</v-card>
```

#### v-list
**Descrizione**: Lista di elementi
**Funzione**: Mostra una lista di elementi con titoli e sottotitoli
```vue
<v-list>
  <v-list-item
    v-for="(item, index) in items"
    :key="index"
    :title="item.title"
    :subtitle="item.subtitle"
  />
</v-list>
```
```

### 6. Interactive Components

#### v-btn
**Descrizione**: Bottone interattivo
**Funzione**: Crea un bottone cliccabile con colori e varianti
```vue
<!-- Attributi principali -->
<v-btn
  color="primary"        <!-- Colore del bottone -->
  variant="text"         <!-- Variant: text, flat, outlined, elevated, tonal, plain -->
  size="normal"          <!-- Size: x-small, small, normal, large, x-large -->
  density="default"      <!-- Density: default, compact, comfortable -->
  prepend-icon="mdi-plus" <!-- Icona precedente -->
  append-icon="mdi-menu" <!-- Icona successiva -->
  :loading="false"      <!-- Stato di caricamento -->
  :disabled="false"     <!-- Disabilitato -->
  :block="false"        <!-- Bottone a larghezza piena -->
  :rounded="false"      <!-- Bordi arrotondati -->
  :elevation="0"        <!-- Elevazione -->
  :ripple="true"        <!-- Effetto onda -->
  :text="false"         <!-- Testo solo -->
  @click="azione"       <!-- Evento click -->
>
  Clicca qui
</v-btn>
```

#### v-dialog
**Descrizione**: Finestra modale
**Funzione**: Mostra contenuto in una finestra sovrapposta
```vue
<!-- Attributi principali -->
<v-dialog
  v-model="dialog"        <!-- Stato del dialog -->
  max-width="500px"       <!-- Larghezza massima -->
  persistent="false"      <!-- Persistente -->
  :fullscreen="false"    <!-- Fullscreen -->
  :transition="dialog-transition" <!-- Transizione -->
  :scrollable="false"    <!-- Scrollabile -->
  :no-click-animation="false" <!-- Disabilita animazione click -->
  :close-on-content-click="false" <!-- Chiude al click sul contenuto -->
>
  <v-card>
    <v-card-title>Titolo</v-card-title>
    <v-card-text>Contenuto</v-card-text>
    <v-card-actions>
      <v-btn @click="dialog = false">Chiudi</v-btn>
    </v-card-actions>
  </v-card>
</v-dialog>
```

#### v-tooltip
**Descrizione**: Scheda informativa
**Funzione**: Mostra informazioni aggiuntive quando si passa sopra un elemento
```vue
<v-tooltip text="Informazioni aggiuntive">
  <template v-slot:activator="{ props }">
    <v-btn v-bind="props">Hover me</v-btn>
  </template>
</v-tooltip>
```
```

### 7. Utility Components

#### v-progress-circular
**Descrizione**: Indicatore di progresso circolare
**Funzione**: Mostra il progresso di un'operazione
```vue
<v-progress-circular
  :model-value="progress"
  color="primary"
/>
```

#### v-alert
**Descrizione**: Notifica di stato
**Funzione**: Mostra messaggi di successo, errore, avviso o informazione
```vue
<v-alert type="success">
  Azione completata con successo!
</v-alert>
```

#### v-snackbar
**Descrizione**: Notifica temporanea
**Funzione**: Mostra brevi messaggi di notifica in basso
```vue
<!-- Attributi principali -->
<v-snackbar
  v-model="snackbar"      <!-- Stato del snackbar -->
  :timeout="5000"        <!-- Timeout in ms -->
  color="success"        <!-- Colore -->
  location="bottom right" <!-- Posizione -->
  :multi-line="false"    <!-- Multi linea -->
  :vertical="false"      <!-- Verticale -->
  :absolute="false"      <!-- Posizione assoluta -->
  :elevation="0"        <!-- Elevazione -->
  :rounded="false"      <!-- Bordi arrotondati -->
>
  Messaggio di notifica
  <template v-slot:actions>
    <v-btn color="pink" variant="text" @click="snackbar = false">
      Close
    </v-btn>
  </template>
</v-snackbar>
```
```

## Caratteristiche Principali

- **Responsive Design**: Tutti i componenti sono nativamente responsive
- **Customizzazione**: Possibilità di personalizzare colori, temi e stili
- **Variante**: Supporto per diverse varianti di ogni componente
- **Direttive**: Direttive utili come v-model, v-if, v-for
- **Eventi**: Gestione degli eventi tramite @click, @change, ecc.

## Best Practices

1. Utilizzare v-container per la gestione del layout
2. Implementare v-navigation-drawer per la navigazione
3. Usare v-card per raggruppare il contenuto
4. Preferire v-select per le selezioni multiple
5. Utilizzare v-dialog per le interazioni modali

### Personalizzazione dei Colori

### Colori Base

#### Rosso (Red)
- `red-light`: #FF6B6B (rosa acceso)
- `red-dark`: #D00000 (rosso scuro)

#### Blu (Blue)
- `blue-light`: #6495ED (blu cielo)
- `blue-dark`: #1E3D59 (blu scuro)

#### Verde (Green)
- `green-light`: #66FF33 (verde brillante)
- `green-dark`: #006400 (verde scuro)

#### Giallo (Yellow)
- `yellow-light`: #FFD700 (oro)
- `yellow-dark`: #DAA520 (oro scuro)

#### Rosa (Pink)
- `pink-light`: #FFB6C1 (rosa pallido)
- `pink-dark`: #E75480 (rosa acceso)

#### Viola (Purple)
- `purple-light`: #9B59B6 (viola brillante)
- `purple-dark`: #5B2C6F (viola scuro)

#### Arancione (Orange)
- `orange-light`: #FFA500 (arancione brillante)
- `orange-dark`: #E67E22 (arancione scuro)

#### Ciano (Cyan)
- `cyan-light`: #00FFFF (ciano brillante)
- `cyan-dark`: #008B8B (ciano scuro)

### Stili di Testo

#### Titoli
- `text-h1`: Titolo principale (48px)
- `text-h2`: Titolo secondario (36px)
- `text-h3`: Titolo terziario (28px)
- `text-h4`: Titolo quaternario (24px)
- `text-h5`: Titolo quinario (20px)
- `text-h6`: Titolo sestario (16px)

#### Testo Base
- `text-body-1`: Testo principale (16px)
- `text-body-2`: Testo secondario (14px)
- `text-subtitle-1`: Sottotitolo (16px)
- `text-subtitle-2`: Sottotitolo secondario (14px)
- `text-caption`: Didascalia (12px)
- `text-overline`: Overline (10px)

#### Stili Speciali
- `text-center`: Allineamento centrale
- `text-left`: Allineamento sinistro
- `text-right`: Allineamento destro
- `text-justify`: Giustificazione
- `text-no-wrap`: No wrap
- `text-truncate`: Troncamento
- `text-lowercase`: Maiuscolo
- `text-uppercase`: Maiuscolo
- `text-capitalize`: Capitalizzazione
- `text-decoration-none`: Nessun decorazione
- `text-decoration-underline`: Sottolineato
- `text-decoration-line-through`: Barrato

#### Spaziature
- `text-xs`: Spaziatura extra small
- `text-sm`: Spaziatura small
- `text-md`: Spaziatura medium
- `text-lg`: Spaziatura large
- `text-xl`: Spaziatura extra large

#### Colori
- `text-primary`: Colore primario
- `text-secondary`: Colore secondario
- `text-accent`: Colore accent
- `text-error`: Colore errore
- `text-info`: Colore informazione
- `text-success`: Colore successo
- `text-warning`: Colore avviso

### Utilizzo dei Colori e Stili
```vue
<!-- Esempio di utilizzo -->
<v-btn color="red-light">Clicca qui</v-btn>
<v-alert type="green-dark">Successo!</v-alert>
<v-card class="blue-light">
  <v-card-title class="text-h5 blue-dark--text">
    Titolo
  </v-card-title>
  <v-card-text class="text-body-1">
    <p class="text-center text-h6">Testo centrato</p>
    <p class="text-body-1">Testo normale</p>
    <p class="text-caption">Didascalia</p>
    <p class="text-overline">Overline</p>
    <p class="text-truncate">Testo troncato</p>
  </v-card-text>
</v-card>
```

## Note Importanti

- Assicurarsi di importare i componenti necessari
- Utilizzare le prop variant per personalizzare l'aspetto
- Gestire lo stato con v-model quando possibile
- Utilizzare le direttive v-slot per il contenuto personalizzato
- Implementare la gestione degli errori con v-alert quando necessario
