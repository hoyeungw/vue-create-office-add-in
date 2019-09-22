<template>
  <v-dialog v-model="dialog" width="500">
    <template #activator="{ on }">
      <v-btn color="red lighten-2" class="ma-1 pa-1" dark v-on="on" @click="onDistinctInitialize">
        distinct
      </v-btn>
    </template>
    <v-stepper v-model="e1">
      <v-stepper-header>
        <v-stepper-step editable :complete="e1 > 1" step="1" @click="onDistinctBackEdit">Range</v-stepper-step>
        <v-divider></v-divider>
        <v-stepper-step :complete="e1 > 2" step="2">Destination</v-stepper-step>
      </v-stepper-header>
      <v-stepper-items>
        <v-stepper-content step="1" class="pa-3">
          <v-card class="ma-1" elevation="0">
            <v-card-title class="headline grey lighten-2 mb-2 subtitle-2" primary-title>
              Choose data source
            </v-card-title>
            <v-card-text>
              Are you selecting <b>{{tempAddress}}</b>?
              You can re-select the range you want to distinct before 'continue'.
            </v-card-text>
          </v-card>
          <v-btn color="primary" class="mx-2" text @click="onDistinctStepOne">
            <v-icon>mdi-check</v-icon>
            Select as source
          </v-btn>
          <v-btn class="mx-2" text @click="dialog = false">
            <v-icon>mdi-close</v-icon>
            Cancel
          </v-btn>
        </v-stepper-content>
        <v-stepper-content step="2" class="pa-3">
          <v-card class="mb-3" elevation="0">
            <v-card-title class="headline grey lighten-2 mb-2 subtitle-2" primary-title>
              Choose destination
            </v-card-title>
            <v-card-text>
              Are you selecting {{targetAddress}} as target?
              You can re-select the range you want to distinct before 'confirm'.
            </v-card-text>
          </v-card>
          <v-btn color="primary" class="mx-2" text @click="onDistinctStepTwo">
            <v-icon>mdi-check</v-icon>
            Put distinct here
          </v-btn>
          <v-btn class="mx-2" text @click="e1 = 1">
            <v-icon>mdi-close</v-icon>
            Cancel
          </v-btn>
        </v-stepper-content>
      </v-stepper-items>
    </v-stepper>
  </v-dialog>
</template>

<script>
  import { deco } from 'xbrief'
  import { Mat as LocalMat, Vec } from '@/functions/algebra'
  import { Mat } from 'veho'

  export default {
    name: 'InstaStepperDistinct',
    props: {
      localLogger: Function
    },
    methods: {
      onDistinctInitialize () {
        window.Excel.run(async (context) => {
          let range = context.workbook.getSelectedRange()
          range.load('address')
          return context.sync()
            .then(() => {
              const addr = range.address;
              `Select range at step one: ${addr}` |> this.localLogger
              this.tempAddress = addr
              this.e1 = 1
            })
            .catch(error => {
              error.name.tag(error.message) |> this.localLogger
            })
        })
      },
      onDistinctBackEdit () {
        window.Excel.run(async (context) => {
          let sheet = context.workbook.worksheets.getActiveWorksheet()
          let range = sheet.getRange(this.tempAddress)
          range.select()
        }).then(() => {
            this.onDistinctInitialize()
          }
        )
      },
      onDistinctStepOne () {
        window.Excel.run(async (context) => {
          let range = context.workbook.getSelectedRange()
          range.load(['getSurroundingRegion', 'getLastColumn', 'getCell', 'address', 'values'])
          let targetCell = range.getSurroundingRegion().getLastColumn().getCell(0, 2)
          targetCell.load('address')
          return context.sync()
            .then(() => {
              const addr = range.address;
              `Select range at step one: ${addr}` |> this.localLogger
              targetCell.select()
              this.matrix = range.values
              this.tempAddress = addr
              this.targetAddress = targetCell.address
              this.e1 = 2
              return addr
            })
            .catch(error => {
              error.name.tag(error.message) |> this.localLogger
            })
        }).then(async (ctx) => {
            'ctx'.tag(deco(ctx)) |>console.log
          }
        )
      },
      onDistinctStepTwo () {
        window.Excel.run(async (context) => {
          // let sheet = context.workbook.worksheets.getActiveWorksheet()
          let target = context.workbook.getSelectedRange()
          target.load('getAbsoluteResizedRange')
          let vector = this.matrix |> LocalMat.matrixToVector |> Vec.distinct
          target = target.getAbsoluteResizedRange(vector.length, 1)
          target.load('address')
          return context.sync()
            .then(() => {
              target.values = [vector] |> Mat.transpose;
              `Pushed to target range: ${target.address}` |> this.localLogger
              // target.select()
              this.e1 = 1
              this.dialog = false
            })
            .catch(error => {
              error.name.tag(error.message) |> this.localLogger
            })
        })
      },
    },
    data: () => ({
      matrix: [[]],
      tempAddress: '',
      targetAddress: '',
      dialog: false,
      e1: 0
    })
  }
</script>

<style scoped>

</style>