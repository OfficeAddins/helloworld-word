var HTMLWebpackPlugin = require('html-webpack-plugin');      
var HTMLWebpackPluginConfig = new HTMLWebpackPlugin({ 
   template: __dirname + '/index.html',      
   filename: 'index.html',      
   inject: 'body'      
}); 

module.exports = { 
  devtool: "source-map", 
  entry: __dirname + '/index.js',   
  output: { 
      filename: 'webpackbundle.js'   
  }, 
  module: { 
    rules: [ 
      { 
        test: /.js?$/, 
        exclude: /node_modules/, 
        use: { 
          loader: 'babel-loader'  
        } 
      } 
    ] 
  }, 
  plugins: [HTMLWebpackPluginConfig] 
}; 