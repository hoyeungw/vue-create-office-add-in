// const VuetifyLoaderPlugin = require('vuetify-loader/lib/plugin')

module.exports = {
  devServer: {
    port: '8084',
    https: true,
  },
  runtimeCompiler: true,
  transpileDependencies: ['vuetify'],
}
