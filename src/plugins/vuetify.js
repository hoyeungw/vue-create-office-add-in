import Vue from 'vue'
import Vuetify from 'vuetify/lib'
import colors from 'vuetify/lib/util/colors'
// import { Ripple } from 'vuetify/lib/directives'

Vue.use(Vuetify, {
  // components: {
  //   VCard,
  //   VRow,
  //   VBtn,
  // },
  // directives: {
  //   Ripple,
  // },
})

export default new Vuetify({
  theme: {
    themes: {
      light: {
        primary: colors.lightBlue.lighten3,
        secondary: '#424242',
        accent: '#82B1FF',
        error: '#FF5252',
        info: '#2196F3',
        success: '#4CAF50',
        warning: '#FFC107',
      },
    },
  },
  icons: {
    iconfont: 'mdi',
  },
})
