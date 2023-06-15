const MiniCssExtractPlugin = require("mini-css-extract-plugin");
const StringReplacePlugin = require("string-replace-webpack-plugin");

module.exports = {
    /*
     * multiple entry points, one per entry
     * the [name] for each is the basename, e.g. some/path/to/thing so we can add .js and .css suffixes
     * the values are the files with their .js6 suffixes retained
     */
    entry: __dirname + '/src/index.js',
    output: {
        path: __dirname + '/docs',
        filename: 'index.js'
    },

    module: {
        rules: [
            /*
             * Plain JS files
             * just kidding; Webpack already does those without any configuration  :)
             * but we do not want to lump them in with ES6 files: they would be third-party and then run through JSHint and we can't waste time linting third-party JS
             */

            /*
             * run .js6 files through Babel + ES2015 via loader; lint and transpile into X.js
             * that's our suffix for ECMAScript 2015 / ES6 files
             */
            {
                test: /\.(js|js6)$/,
                use: [
                    { loader: 'babel-loader', options: { presets: ['@babel/env'] } }
                ],
                exclude: /node_modules/
            },

            /*
             * CSS files and also SASS-to-CSS all go into one bundled X.css
             */
            {
                test: /\.css$/,
                    use: [
                        { loader: MiniCssExtractPlugin.loader, options: { publicPath: './docs', } },
                        { loader: 'css-loader', options: { sourceMap: true } }
                    ]
            },
            {
                test: /\.scss$/,
                    use: [
                        { loader: MiniCssExtractPlugin.loader, options: { publicPath: './docs', } },
                        { loader: 'css-loader', options: { sourceMap: true} },
                        { loader: 'sass-loader', options: { sourceMap:true } },
                    ]
            },

            /*
             * HTML Files
             * replace [hash] entities in the .src.html to generate .html
             * typically used on .js and .css filenames to include a random hash for cache-busting
             * though could be used to cache-bust nearly anything such as images
             * tip: HTML file basenames (like any) should be fairly minimal: letters and numbers, - _ . characters
             */
            {
                test: /\.html$/,
                use: [
                    { loader: 'file-loader', options: { name: '[name].html' } },
                    {
                        loader: StringReplacePlugin.replace({
                        replacements: [
                            {
                                pattern: /\[hash\]/g,
                                replacement: function (match, p1, offset, string) {
                                    const randomhash = new Date().toString().replace(/\W/g, '_');
                                    return randomhash;
                                }
                            },
                        ]})
                    }
                ]
            },

            /*
             * Images like in CSS
             */
            {
                test: /\.(svg|gif|jpg|jpeg|png)$/,
                loader: 'url-loader',
                options: { 
                    limit: 1000, // Convert images < 1kb to base64 strings
                    name: 'images/[hash]-[name].[ext]'
                }

            },

            /*
             * Files to ignore
             * Notably from CSS, e.g. background-image SVG, PNGs, JPEGs, fonts, ...
             * we do not need them processed; our stylesheets etc. will point to them in their proper place
             * webpack scans the HTML files and will throw a fit if we don't account for every single file it finds
             */
            {
                test: /\.(woff|woff2|ttf|eot)$/,
                loader: 'ignore-loader'
            }
        ]
    },


    /*
     * enable source maps, applicable to both JS and CSS
     */
    devtool: "nosources-source-map",

    /*
     * plugins for the above
     */
    plugins: [
        // CSS output from the CSS + LESS handlers above
        new MiniCssExtractPlugin({
            filename: 'index.css',
        }),
        // for doing string replacements on files
        new StringReplacePlugin(),
    ],

    /*
     * plugins for the above
     */
    devServer: {
        contentBase: './docs/',
        port: 8182,
        compress: true,
        watchContentBase: true, // live reload
    }
};
