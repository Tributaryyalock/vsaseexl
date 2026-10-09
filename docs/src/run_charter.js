const { charter } = require('./charter_lib');
const [,, mod, out] = process.argv;
charter(require('./' + mod)(out));
