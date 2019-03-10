const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyPlugin = require('copy-webpack-plugin');

module.exports = {
    entry: {
        polyfill: 'babel-polyfill',
        app: './src/index.js',
        'function-file': './function-file/function-file.js'
    },
    module: {
        rules: [
            {
                test: /\.js$/,
                exclude: /node_modules/,
                use: 'babel-loader'
            },
            {
                test: /\.html$/,
                exclude: /node_modules/,
                use: 'html-loader'
            },
            {
                test: /\.(png|jpg|jpeg|gif)$/,
                use: 'file-loader'
            }
        ]
    },
    plugins: [
        new HtmlWebpackPlugin({
            template: './index.html',
            chunks: ['polyfill', 'app']
        }),
        new HtmlWebpackPlugin({
            template: './function-file/function-file.html',
            filename: 'function-file/function-file.html',
            chunks: ['function-file']
        }),
        new CopyPlugin([
            './app.css',
            './resource.html',
            './support.html',
            './ios_manifest.xml',
            './web_manifest.xml',
            { from: './assets', to: "assets" },
            { from: './node_modules/jquery/dist/jquery.min.js', to: "node_modules/jquery/dist/jquery.min.js" },
            { from: './node_modules/office-ui-fabric-js/dist/js', to: "node_modules/office-ui-fabric-js/dist/js" },
            { from: './node_modules/office-ui-fabric-js/dist/css', to: "node_modules/office-ui-fabric-js/dist/css" }
        ], {
            to: './dist'
        })
    ]
};