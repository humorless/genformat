# genformat

FIXME: description

## Installation
1. install java
2. install leiningen
3. execute the command to build executable jar: `genformat-1.0.0-standalone.jar`
   `lein uberjar` 
4. move the file `target/uberjar/genformat-1.0.0-standalone.jar` to the right directory.

## Usage

Generating sub ticket from tsv files
```
    $ java -jar genformat-1.0.0-standalone.jar :tsv $TEH_TSV_INPUT_FILE_PATH
```

Generating sub ticket from database

Note: using this option, you need to config the file `config.edn`
```
    $ java -jar genformat-1.0.0-standalone.jar :db
```
## Examples


### Bugs


## License

Copyright © 2022 Laurence Chen from [REPLWARE](https://replware.dev)

This program and the accompanying materials are made available under the
terms of the Eclipse Public License 2.0 which is available at
http://www.eclipse.org/legal/epl-2.0.

This Source Code may also be made available under the following Secondary
Licenses when the conditions for such availability set forth in the Eclipse
Public License, v. 2.0 are satisfied: GNU General Public License as published by
the Free Software Foundation, either version 2 of the License, or (at your
option) any later version, with the GNU Classpath Exception which is available
at https://www.gnu.org/software/classpath/license.html.
