<template>
  <v-app id="app">
    <v-navigation-drawer app>
    </v-navigation-drawer>
    <v-app-bar color="cyan lighten-1" class="elevation-0" dense dark app>
      <v-app-bar-nav-icon></v-app-bar-nav-icon>
      <v-toolbar-title>CrosTab</v-toolbar-title>
      <div class="flex-grow-1"></div>
      <v-btn icon>
        <v-icon>mdi-share-variant</v-icon>
      </v-btn>
      <v-btn icon>
        <v-icon>mdi-magnify</v-icon>
      </v-btn>
      <v-menu left bottom>
        <template #activator="{ on }">
          <v-btn icon v-on="on">
            <v-icon>mdi-dots-vertical</v-icon>
          </v-btn>
        </template>
        <v-list>
          <v-list-item v-for="n in 5" :key="n" @click="() => {}">
            <v-list-item-title>Option {{ n }}</v-list-item-title>
          </v-list-item>
        </v-list>
      </v-menu>
    </v-app-bar>
    <v-content>
      <v-container fluid>
        <v-expansion-panels v-model="panel" focusable accordion multiple>
          <v-expansion-panel id="done-by-click">
            <v-expansion-panel-header>Done by Click</v-expansion-panel-header>
            <v-expansion-panel-content>
              <v-row class="d-flex flex-row">
                <insta-stepper-distinct :local-logger="pushToSnackbar"></insta-stepper-distinct>
                <v-tooltip id="button-get-selected" cols="4" xs="2" bottom>
                  <template #activator="{ on }">
                    <v-btn class="ma-1 pa-1" elevation=1 color="primary" dark v-on="on"
                           @click="onGetSelectedRange">
                      selected
                    </v-btn>
                  </template>
                  <span>Tooltip</span>
                </v-tooltip>
                <v-btn class="ma-1 pa-1" elevation=1 cols="4" xs="2" @click="onGetSurrounding">surround</v-btn>
              </v-row>
            </v-expansion-panel-content>
          </v-expansion-panel>
          <v-expansion-panel id="cros-tab">
            <v-expansion-panel-header>CrosTab</v-expansion-panel-header>
            <v-expansion-panel-content>
              <v-row>
                <!--                <v-col cols="12">-->
                <v-combobox
                    v-model="select"
                    :items="fields"
                    chips
                    label="Side"
                ></v-combobox>
                <v-combobox
                    v-model="select"
                    :items="fields"
                    chips
                    label="Banner"
                ></v-combobox>
                <v-combobox
                    v-model="selects"
                    :items="aggregates"
                    chips
                    multiple
                    label="Aggregates"
                ></v-combobox>
                <v-combobox
                    v-model="selects"
                    :items="filters"
                    chips
                    multiple
                    label="Filters"
                ></v-combobox>
                <!--                </v-col>-->
                <v-btn color="primary" class="elevation-0" block dark>
                  Cross table
                </v-btn>
              </v-row>
            </v-expansion-panel-content>
          </v-expansion-panel>
        </v-expansion-panels>
        <div class="text-left my-4">
          <v-subheader>
            Thanks for using.
          </v-subheader>
          <v-subheader v-if="address">
            Last selected: {{address}}
          </v-subheader>
        </div>
        <v-snackbar v-model="snackbar">
          {{ message }}
          <v-btn color="warning" icon light small @click="snackbar = false">
            <v-icon>mdi-close</v-icon>
          </v-btn>
        </v-snackbar>
      </v-container>
    </v-content>
    <v-footer color="cyan lighten-1" app>
      <div class="flex-grow-1"></div>
      <span class="white--text">&copy; Leagyun Tech 2019</span>
    </v-footer>
  </v-app>
</template>

<script>
  import { GP } from 'elprimero'
  import InstaStepperDistinct from '@/components/InstaStepperDistinct.vue'

  GP.now().tag('App.vue') |> console.log
  export default {
    name: 'App',
    methods: {
      pushToSnackbar (message = '') {
        if (message) this.message = message
        this.snackbar = true
        setTimeout(() => {this.snackbar = false}, 3000)
      },
      onGetSelectedRange () {
        window.Excel.run(async (context) => {
          let range = context.workbook.getSelectedRange()
          range.load('address')
          return context.sync()
            .then(() => {
              const addr = range.address;
              `Select range: ${addr}` |> this.pushToSnackbar
            })
            .catch(error => {
              error.name.tag(error.message) |> this.pushToSnackbar
            })
        })
      },
      onGetSurrounding () {
        window.Excel.run(async (context) => {
          let range = context.workbook.getSelectedRange()
          range.load('getSurroundingRegion')
          let surroundingRegion = range.getSurroundingRegion()
          surroundingRegion.load('address')
          await context.sync()
            .then(() => {
              `Select surrounding region: ${surroundingRegion.address}` |> this.pushToSnackbar
              surroundingRegion.select()
              this.address = surroundingRegion.address
            })
            .catch(error => {
              console.log(`Error: ${error}`)
            })
        })
      },
      onShowDialog () {
        window.Excel.run(async (context) => {
          window.Office.context.ui.displayDialogAsync('https://baidu.com', {
              promptBeforeOpen: false,
              height: 30,
              width: 20
            },
            function (asyncResult) {
              console.log(asyncResult)
              // dialog = asyncResult.value;
              // dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
            }
          )
          await context.sync()
            .then(() => {
            })
            .catch(error => {
              console.log(`Error: ${error}`)
            })
        })
      },
      onGetSelected () {
        window.Excel.run(function (context) {
          const range = context.workbook.getSelectedRange()
          range.load('address')
          return context.sync().then(function () {
            const addr = range.address;
            (`The address of the selected range is "${addr}"`) |> this.pushToSnackbar
            this.tempAddress = addr
          })
        })
      },
      onSaveSetting () {
        // localStorage.timeStamp = GP.present()
      },
      deciferCrosTab () {

      },
      onChooseRange () {
        window.Excel.run(async (context) => {
          let range = context.workbook.getSelectedRange()
          range.load('address')
          let addr
          await context.sync()
            .then(() => {
              addr = range.address
            })
            .then(() => {
              console.log(addr)
              // this.address = range.address;
              // `Selected region: ${this.address}` |> this.pushToSnackbar
              // localStorage.address = this.address
            })
            .catch(error => {
              error.name.tag(error.message) |> this.pushToSnackbar
            })
        })
        this.dialog = false
      }
    },
    mounted () {
      if (localStorage.address) {
        this.address = localStorage.address
      }
      'mounted'.tag('localStorage.address').tag(localStorage.address).wL()
    },
    // watch: {
    //   name (val) {
    //     localStorage.name = val
    //   }
    // },
    data: () => ({
      panel: [0, 1],
      address: '',
      dialog: false,
      tags: [
        'distinct',
        'asphalt',
        'separate',
        'concate',
      ],
      snackbar: false,
      message: 'Hello, I\'m a snackbar 2.',
      select: 'Programming',
      selects: ['Programming'],
      fields: [
        'City',
        'Gender',
        'Taste',
        'ToM',
        'Rate'
      ],
      aggregates: [
        'Taste',
        'ToM',
        'Rate'
      ],
      filters: [
        'City',
        'Gender'
      ],
      filterInstances: {
        'City': c => ['Shanghai', 'Beijing'].includes(c),
        'Gender': g => g === 'Female'
      }
    }),
    components: {
      InstaStepperDistinct
    }
  }
</script>

<style>
</style>
