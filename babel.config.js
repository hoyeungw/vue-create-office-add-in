module.exports = {
  presets: [
    // ['@vue/app'],
    ['@babel/preset-env'],
  ],
  plugins: [
    ['@babel/plugin-proposal-pipeline-operator', { 'proposal': 'minimal' }]
  ]
}
