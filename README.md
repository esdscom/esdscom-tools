# eSDScom Tools

Currently, this repository contains just one tool.

## eSDScom XSD to SAMM converter

### Purpose

For use of the eSDScom data structure in the context of Digital Product Passports (DPP) and Catena-X, an automotive industry project for DPPs, Data must be sent as JSON instead of XML. Such JSON data is specified as a turtle ontology (TTL file) in SAMM format.

The eSDScom team wanted to maintain just one specification in order to guarantee consistency between the XSD and TTL flavour of eSDScom. Thus an automatic conversion from XSD to TTL was the method of choice. However, there are rules and properties specific to SAMM, and Catena-X, such as (and surely not limited to):

* TTL files cannot use the same property name (= XSD element) with different types in different places
* Naming conventions apply (lower / upper case first letter, camel case, usable characters)
* Decriptions and example values are mandatory in SAMM

The conversion script takes care of this. It uses existing annotations als descriptions, converts appinfo annotations into samm:see to designate recommended subsets of eSDScom, and converts appinfos marked by `<esdscom:example>` into SAMM example values.

### Limitations

While generally applicable to every XSD to SAMM conversion, the author did not bother about XSD features not used in eSDScom, and the script may refer to eSDScom in various places. Review this upon forking for your own purpose.

We would love to make this a generally usable XSD to SAMM converter. If you apply changes that do not break usability for eSDScom, please issue a pull request to make them available for others.

### Links

https://www.esdscom.eu/
https://eclipse-esmf.github.io/samm-specification/snapshot/index.html
https://docs.bosch-semantic-stack.com/oss/aspect-model-editor.html
https://catenax-ev.github.io/
