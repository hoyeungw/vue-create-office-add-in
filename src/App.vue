<template>
  <div id="app">
    <v-app-bar color="deep-purple accent-4" dense dark>
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
        <template v-slot:activator="{ on }">
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
    <v-expansion-panels>
      <v-expansion-panel id="done-by-click">
        <v-expansion-panel-header>Done by Click</v-expansion-panel-header>
        <v-expansion-panel-content>
          <v-row>
            <v-col cols="12">
              <v-row justify="center">
                <v-col cols="4" xs="2">
                  <v-btn class="mx-2 caption" elevation=3 @click="onGetSelectedRange">selected</v-btn>
                </v-col>
                <v-col cols="4" md="2">
                  <v-btn class="mx-2" elevation=3 @click="onGetSurrounding">surround</v-btn>
                </v-col>
                <v-col cols="4" md="2">
                  <v-btn class="mx-2" elevation=3 @click="onDistinct">distinct</v-btn>
                </v-col>
              </v-row>
            </v-col>
          </v-row>
        </v-expansion-panel-content>
      </v-expansion-panel>
      <v-expansion-panel id="cros-tab">
        <v-expansion-panel-header>CrosTab</v-expansion-panel-header>
        <v-expansion-panel-content>
          <v-row>
            <v-col cols="12">
              <v-row justify="center">
                <v-col cols="4" xs="2">
                  <v-btn class="mx-4 " elevation=3 @click="onShowDialog">dialog</v-btn>
                </v-col>
                <v-col cols="4" md="2">
                  <v-btn class="mx-2" elevation=3 @click="onGetSurrounding">surround</v-btn>
                </v-col>
                <v-col cols="4" md="2">
                  <v-btn class="mx-2" elevation=3 @click="onDistinct">distinct</v-btn>
                </v-col>
              </v-row>
            </v-col>
          </v-row>
        </v-expansion-panel-content>
      </v-expansion-panel>
    </v-expansion-panels>
    <v-container fluid>
      <div class="text-center">
        <p>Choose the button below to set the color of the selected range to green.</p>
        <v-subheader>Last selected: {{address}}</v-subheader>
        <!--        <br/>-->
        <!--        <h3>Try it out</h3>-->
        <v-subheader class="headline">Try it out</v-subheader>
        <v-dialog v-model="dialog" width="500">
          <template v-slot:activator="{ on }">
            <v-btn color="red lighten-2" dark v-on="on">
              choose range
            </v-btn>
          </template>
          <v-card>
            <v-card-title class="headline grey lighten-2" primary-title>
              choose range
            </v-card-title>
            <v-card-text>
              Please choose a target range in the worksheet. Then return and click 'confirm'.
            </v-card-text>
            <v-divider></v-divider>
            <v-card-actions>
              <div class="flex-grow-1"></div>
              <v-btn color="primary" text @click="onChooseRange">
                confirm
              </v-btn>
            </v-card-actions>
          </v-card>
        </v-dialog>
      </div>
      <v-card max-width="400" class="mx-auto">
        <v-card-text>
          <v-chip-group multiple column active-class="primary--text">
            <v-chip label color="primary" v-for="tag in tags" :key="tag">
              {{ tag }}
            </v-chip>
          </v-chip-group>
        </v-card-text>
      </v-card>
      <div class="text-center ma-2">
        <v-subheader class="headline">Snackbar</v-subheader>
        <!--        <v-btn dark @click="onSnackbarClicked()">Open Snackbar</v-btn>-->
        <v-snackbar v-model="snackbar">
          {{ message }}
          <v-btn color="pink" text @click="snackbar = false">
            Close
          </v-btn>
        </v-snackbar>
      </div>
    </v-container>
  </div>
</template>

<script>
  import { VecX, StrX } from 'xbrief'

  StrX.wL('App.vue')
  export default {
    name: 'App',
    methods: {
      onSnackbarClicked (message = '') {
        if (message) this.message = message
        this.snackbar = true
        setTimeout(() => {this.snackbar = false}, 3000)
      },
      onGetSelectedRange () {
        window.Excel.run(function (context) {
          let range = context.workbook.getSelectedRange()
          range.load('address')
          return context.sync()
            .then(function () {
              console.log(`The address of the selected range is "${range.address}"`)
            })
            .catch(error => {
              error.name.tag(error.message).wL()
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
              `Select surrounding region: ${surroundingRegion.address}` |> this.onSnackbarClicked
              surroundingRegion.select()
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
      onDistinct () {
        const matrixToVector = (matrix) => {
          const arr = []
          for (let row of matrix) for (let x of row) arr.push(x)
          return arr
        }
        window.Excel.run(async (context) => {
          let range = context.workbook.getSelectedRange()
          range.load('getSurroundingRegion')
          let surroundingRegion = range.getSurroundingRegion()
          surroundingRegion.load('address')
          surroundingRegion.load('values')
          await context.sync().then(() => {
            const vec = matrixToVector(surroundingRegion.values)
            VecX.hBrief(vec).wL();
            `The address of the surrounding region is "${surroundingRegion.address}"`.wL()
            // range.getCell(0, 0).values = [[range.getSpillingToRange().address]]
            surroundingRegion.select()
          })
          // .catch(error => {
          //   console.log(`Error: ${error}`)
          // })
        })
      },
      onGetSelected () {
        window.Excel.run(function (context) {
          const range = context.workbook.getSelectedRange()
          range.load('address')
          return context.sync().then(function () {
            console.log(`The address of the selected range is "${range.address}"`)
          })
        })
      },
      onSaveSetting () {
        localStorage.timeStamp = GP.present()
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
              // `Selected region: ${this.address}` |> this.onSnackbarClicked
              // localStorage.address = this.address
            })
            .catch(error => {
              error.name.tag(error.message) |> this.onSnackbarClicked
            })
        })
        this.dialog = false
      }
    },
    mounted () {
      if (localStorage.address) {
        this.address = localStorage.address
      }
    },
    // watch: {
    //   name (val) {
    //     localStorage.name = val
    //   }
    // },
    data: () => ({
      address: '',
      dialog: false,
      tags: [
        'distinct',
        'asphalt',
        'seperate',
        'concate',
      ],
      snackbar: false,
      message: 'Hello, I\'m a snackbar'
    }),
  }
</script>

<style>
</style>
