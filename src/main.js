import 'babel-polyfill'
import 'roboto-fontface/css/roboto/roboto-fontface.css'
import '@mdi/font/css/materialdesignicons.css'
import Vue from 'vue'
import App from './App.vue'
import vuetify from './plugins/vuetify'
import { Str } from 'xbrief'

Vue.config.productionTip = false

// const Office = window.Office
// Office.initialize = () => {
//   new Vue({
//     vuetify,
//     render: h => h(App),
//   }).$mount('#app')
// }

// new Vue({
//   vuetify,
//   render: h => h(App)
// }).$mount('#app')
Str.wL('now it\'s time to load main.js')

const Office = window.Office
Office.initialize = () => {
  new Vue({
    vuetify,
    el: '#app',
    components: { App },
    template: '<App/>',
  })
}
