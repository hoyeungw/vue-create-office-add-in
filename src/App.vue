<template>
  <div id="app">
    <div id="content">
      <div id="content-header">
        <div class="padding">
          <h1>Welcome</h1>
        </div>
      </div>
      <div id="content-main">
        <div class="padding">
          <p>Choose the button below to set the color of the selected range to green.</p>
          <br/>
          <h3>Try it out</h3>
          <v-btn @click="onAbsoluteResize">Resize</v-btn>
          <v-btn @click="onSpillRelated">Spill</v-btn>
          <v-btn @click="onGetSelected">Get Selected</v-btn>
        </div>
      </div>
    </div>
  </div>
</template>

<script>
  import { absoluteResize } from './functions/absoluteResize'
  // import { Str } from 'xbrief'

  export default {
    name: 'App',
    methods: {
      onAbsoluteResize () {
        window.Excel.run(absoluteResize)
      },
      onSpillRelated () {
        window.Excel.run(async (context) => {
          let range = context.workbook.getSelectedRange()
          range.load('address')
          range.load('getSurroundingRegion')
          // range.format.fill.color = 'green'
          // const matrix = Mat.ini(3, 3, (i, j) => i + j + 1)
          // MatX.xBrief(matrix).wL()
          let spilledRange = range.getSurroundingRegion()
          spilledRange.load('address')
          await context.sync().then(() => {
            console.log(range.address)
            console.log(spilledRange.address)
            // 'range.address'.tag(range.address).wL()
            // range.getCell(0, 0).values = [[range.getSpillingToRange().address]]
          }).catch(error => {
            console.log(`Error: ${error}`)
          })
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
    },
  }
</script>

<style>
  #content-header {
    background: #2a8dd4;
    color: #fff;
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
    height: 80px;
    overflow: hidden;
  }

  #content-main {
    background: #fff;
    position: fixed;
    top: 80px;
    left: 0;
    right: 0;
    bottom: 0;
    overflow: auto;
  }

  .padding {
    padding: 15px;
  }
</style>
