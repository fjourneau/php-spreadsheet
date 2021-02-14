# php-spreadsheet

## Install via composer

Not yet available on packagist.org. Waiting stable version.
 
Add in composer.json :
```` json
    "repositories": [
        {
            "type": "vcs",
            "url": "https://github.com/fjourneau/php-spreadsheet.git"
        }
    ],
    "require": {
        "fjourneau/spreadsheet": "dev-master"
    }
    
````

If not mentionned, add :
````json
    "minimum-stability": "dev"
````


Then run ``composer install`` or ``composer update``.

_Notice than Slim 3 is required and not added in composer dependancies._