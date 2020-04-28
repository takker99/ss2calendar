module.exports = {
    extends: [
        'plugin:@typescript-eslint/recommended',
        'prettier',
        'prettier/@typescript-eslint',
    ],
    plugins: [
        '@typescript-eslint',
        'prettier'
    ],
    parser: '@typescript-eslint/parser',
    parserOptions: {
        ecmaVersion: '11',
        sourceType: 'module',
        project: './tsconfig.json'
    },
    rules: {
        'prettier/prettier': [
            'error',
            module.exports = {
                printWidth: 80,
                semi: false,
                singleQuote: true,
                trailingComma: 'es5',
                tabWidth: 4,
            }
        ],
        "brace-style": [
            "allman",
            {"allowSingleLine": "true"},
        ]
    }
}
