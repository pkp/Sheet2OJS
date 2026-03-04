<?php

declare(strict_types=1);

/**
 * Build-time schema order map generator.
 *
 * Usage:
 *   php generate_schema_order_map.php
 *
 * Output:
 *   schema_order_map.php (in the same folder)
 */

final class SchemaOrderMapGenerator
{
    private string $rootSchemaPath;
    private string $baseDir;

    /** @var array<string, DOMDocument> */
    private array $documents = [];

    /** @var array<string, DOMElement> */
    private array $globalElements = [];

    /** @var array<string, DOMElement> */
    private array $complexTypes = [];

    /** @var array<string, list<string>> */
    private array $substitutionChildren = [];

    /** @var array<string, bool> */
    private array $localeTypeCache = [];

    public function __construct(string $rootSchemaPath)
    {
        $this->rootSchemaPath = realpath($rootSchemaPath) ?: $rootSchemaPath;
        $this->baseDir = dirname($this->rootSchemaPath);
    }

    public function run(string $outputPath): void
    {
        $this->loadSchemaRecursive($this->rootSchemaPath);
        $this->buildIndexes();

        $map = [
            'articleElementOrder' => $this->getElementOrderByComplexType('article', 'submission'),
            'publicationElementOrder' => $this->getElementOrderByComplexType('publication', 'publication'),
            'authorElementOrder' => $this->getElementOrderByComplexType('author', 'author'),
            'submissionFileElementOrder' => $this->getElementOrderByComplexType('submission_file', 'submission_file'),
            'issueElementOrder' => $this->getElementOrderByComplexType('issue', 'issue'),
            'issueIdentificationElementOrder' => $this->getElementOrderByElementName('issue_identification'),
            'coverImageElementOrder' => $this->getElementOrderByElementName('cover'),
            'elementsWithLocale' => $this->getElementsWithLocaleAttribute(),
        ];

        $php = "<?php\n\nreturn " . var_export($map, true) . ";\n";
        file_put_contents($outputPath, $php);

        echo "Generated order map: {$outputPath}\n";
    }

    private function loadSchemaRecursive(string $path): void
    {
        $realPath = realpath($path) ?: $path;
        if (isset($this->documents[$realPath])) {
            return;
        }

        $dom = new DOMDocument();
        $dom->preserveWhiteSpace = false;
        $dom->load($realPath);
        $this->documents[$realPath] = $dom;

        $xpath = new DOMXPath($dom);
        $includeNodes = $xpath->query("/*[local-name()='schema']/*[local-name()='include']");
        if (!$includeNodes) {
            return;
        }

        foreach ($includeNodes as $includeNode) {
            if (!$includeNode instanceof DOMElement) {
                continue;
            }
            $schemaLocation = $includeNode->getAttribute('schemaLocation');
            if ($schemaLocation === '') {
                continue;
            }
            $included = dirname($realPath) . DIRECTORY_SEPARATOR . $schemaLocation;
            $this->loadSchemaRecursive($included);
        }
    }

    private function buildIndexes(): void
    {
        foreach ($this->documents as $doc) {
            $xpath = new DOMXPath($doc);

            $complexTypeNodes = $xpath->query("/*[local-name()='schema']/*[local-name()='complexType'][@name]");
            if ($complexTypeNodes) {
                foreach ($complexTypeNodes as $node) {
                    if (!$node instanceof DOMElement) {
                        continue;
                    }
                    $name = $node->getAttribute('name');
                    if ($name !== '' && !isset($this->complexTypes[$name])) {
                        $this->complexTypes[$name] = $node;
                    }
                }
            }

            $elementNodes = $xpath->query("/*[local-name()='schema']/*[local-name()='element'][@name]");
            if ($elementNodes) {
                foreach ($elementNodes as $node) {
                    if (!$node instanceof DOMElement) {
                        continue;
                    }
                    $name = $node->getAttribute('name');
                    if ($name !== '' && !isset($this->globalElements[$name])) {
                        $this->globalElements[$name] = $node;
                    }
                }
            }
        }

        foreach ($this->globalElements as $elementName => $element) {
            $subGroup = $this->stripPrefix($element->getAttribute('substitutionGroup'));
            if ($subGroup === '') {
                continue;
            }
            $this->substitutionChildren[$subGroup] ??= [];
            $this->substitutionChildren[$subGroup][] = $elementName;
        }

        foreach ($this->substitutionChildren as &$children) {
            $children = array_values(array_unique($children));
            sort($children);
        }
    }

    /**
     * Prefer element-backed order when available, else fallback to complex type directly.
     *
     * @return list<string>
     */
    private function getElementOrderByComplexType(string $elementName, string $complexTypeName): array
    {
        if (isset($this->globalElements[$elementName])) {
            $order = $this->getElementOrderByElementName($elementName);
            if ($order !== []) {
                return $order;
            }
        }

        return $this->extractComplexTypeOrder($complexTypeName);
    }

    /**
     * @return list<string>
     */
    private function getElementOrderByElementName(string $elementName): array
    {
        $element = $this->globalElements[$elementName] ?? null;
        if (!$element) {
            return [];
        }

        $typeName = $this->stripPrefix($element->getAttribute('type'));
        if ($typeName !== '') {
            return $this->extractComplexTypeOrder($typeName);
        }

        $inlineComplexType = $this->findDirectChild($element, 'complexType');
        if (!$inlineComplexType) {
            return [];
        }

        return $this->extractOrderFromComplexTypeNode($inlineComplexType);
    }

    /**
     * @return list<string>
     */
    private function extractComplexTypeOrder(string $complexTypeName): array
    {
        $complexType = $this->complexTypes[$complexTypeName] ?? null;
        if (!$complexType) {
            return [];
        }

        return $this->extractOrderFromComplexTypeNode($complexType);
    }

    /**
     * @return list<string>
     */
    private function extractOrderFromComplexTypeNode(DOMElement $complexType): array
    {
        $order = [];

        $extension = $this->findFirstDescendantPath($complexType, ['complexContent', 'extension']);
        if ($extension) {
            $baseType = $this->stripPrefix($extension->getAttribute('base'));
            if ($baseType !== '') {
                $order = array_merge($order, $this->extractComplexTypeOrder($baseType));
            }

            $sequence = $this->findDirectChild($extension, 'sequence');
            if ($sequence) {
                $order = array_merge($order, $this->extractOrderFromCompositor($sequence));
            }
        }

        $directSequence = $this->findDirectChild($complexType, 'sequence');
        if ($directSequence) {
            $order = array_merge($order, $this->extractOrderFromCompositor($directSequence));
        }

        return $this->uniquePreserveOrder($order);
    }

    /**
     * @return list<string>
     */
    private function extractOrderFromCompositor(DOMElement $node): array
    {
        $result = [];

        foreach ($node->childNodes as $child) {
            if (!$child instanceof DOMElement) {
                continue;
            }

            $local = $child->localName;
            if ($local === 'element') {
                $result = array_merge($result, $this->resolveElementDeclaration($child));
                continue;
            }

            if ($local === 'choice' || $local === 'sequence' || $local === 'all') {
                $result = array_merge($result, $this->extractOrderFromCompositor($child));
            }
        }

        return $result;
    }

    /**
     * @return list<string>
     */
    private function resolveElementDeclaration(DOMElement $elementDecl): array
    {
        $ref = $this->stripPrefix($elementDecl->getAttribute('ref'));
        if ($ref !== '') {
            return $this->resolveReferenceToConcreteElements($ref);
        }

        $name = $elementDecl->getAttribute('name');
        if ($name !== '') {
            return [$name];
        }

        return [];
    }

    /**
     * Expand references through substitution groups to concrete element names.
     *
     * @return list<string>
     */
    private function resolveReferenceToConcreteElements(string $refName): array
    {
        $children = $this->substitutionChildren[$refName] ?? [];
        if ($children === []) {
            return [$refName];
        }

        $resolved = [];
        foreach ($children as $childName) {
            $resolved[] = $childName;
            foreach ($this->resolveReferenceToConcreteElements($childName) as $grandChild) {
                if ($grandChild !== $childName) {
                    $resolved[] = $grandChild;
                }
            }
        }

        return $this->uniquePreserveOrder($resolved);
    }

    /**
     * @return list<string>
     */
    private function getElementsWithLocaleAttribute(): array
    {
        $result = [];

        foreach ($this->globalElements as $name => $element) {
            if ($this->elementHasLocale($element)) {
                $result[] = $name;
            }
        }

        foreach ($this->documents as $doc) {
            $xpath = new DOMXPath($doc);
            $nodes = $xpath->query("//*[local-name()='element'][@name and not(parent::*[local-name()='schema'])]");
            if (!$nodes) {
                continue;
            }
            foreach ($nodes as $node) {
                if (!$node instanceof DOMElement) {
                    continue;
                }
                $name = $node->getAttribute('name');
                if ($name === '') {
                    continue;
                }
                if ($this->elementHasLocale($node)) {
                    $result[] = $name;
                }
            }
        }

        $result = array_values(array_unique($result));
        sort($result);
        return $result;
    }

    private function elementHasLocale(DOMElement $element): bool
    {
        $typeName = $this->stripPrefix($element->getAttribute('type'));
        if ($typeName !== '' && $this->typeHasLocale($typeName)) {
            return true;
        }

        $inlineComplexType = $this->findDirectChild($element, 'complexType');
        if ($inlineComplexType && $this->complexTypeNodeHasLocale($inlineComplexType)) {
            return true;
        }

        return false;
    }

    private function typeHasLocale(string $typeName): bool
    {
        if (array_key_exists($typeName, $this->localeTypeCache)) {
            return $this->localeTypeCache[$typeName];
        }

        $this->localeTypeCache[$typeName] = false;

        $typeNode = $this->complexTypes[$typeName] ?? null;
        if (!$typeNode) {
            return false;
        }

        $hasLocale = $this->complexTypeNodeHasLocale($typeNode);
        $this->localeTypeCache[$typeName] = $hasLocale;
        return $hasLocale;
    }

    private function complexTypeNodeHasLocale(DOMElement $complexType): bool
    {
        foreach ($complexType->childNodes as $child) {
            if (!$child instanceof DOMElement) {
                continue;
            }

            if ($child->localName === 'attribute' && $child->getAttribute('name') === 'locale') {
                return true;
            }

            if ($child->localName === 'complexContent') {
                $extension = $this->findDirectChild($child, 'extension');
                if ($extension) {
                    $baseType = $this->stripPrefix($extension->getAttribute('base'));
                    if ($baseType !== '' && $this->typeHasLocale($baseType)) {
                        return true;
                    }

                    foreach ($extension->childNodes as $extChild) {
                        if ($extChild instanceof DOMElement && $extChild->localName === 'attribute' && $extChild->getAttribute('name') === 'locale') {
                            return true;
                        }
                    }
                }
            }

            if ($child->localName === 'simpleContent') {
                $extension = $this->findDirectChild($child, 'extension');
                if ($extension) {
                    foreach ($extension->childNodes as $extChild) {
                        if ($extChild instanceof DOMElement && $extChild->localName === 'attribute' && $extChild->getAttribute('name') === 'locale') {
                            return true;
                        }
                    }
                }
            }
        }

        return false;
    }

    private function stripPrefix(string $qName): string
    {
        if ($qName === '') {
            return '';
        }
        $parts = explode(':', $qName, 2);
        return $parts[count($parts) - 1];
    }

    private function findDirectChild(DOMElement $parent, string $localName): ?DOMElement
    {
        foreach ($parent->childNodes as $child) {
            if ($child instanceof DOMElement && $child->localName === $localName) {
                return $child;
            }
        }
        return null;
    }

    /**
     * @param list<string> $path
     */
    private function findFirstDescendantPath(DOMElement $root, array $path): ?DOMElement
    {
        $current = $root;
        foreach ($path as $segment) {
            $next = $this->findDirectChild($current, $segment);
            if (!$next) {
                return null;
            }
            $current = $next;
        }
        return $current;
    }

    /**
     * @param list<string> $values
     * @return list<string>
     */
    private function uniquePreserveOrder(array $values): array
    {
        $seen = [];
        $out = [];

        foreach ($values as $value) {
            if ($value === '' || isset($seen[$value])) {
                continue;
            }
            $seen[$value] = true;
            $out[] = $value;
        }

        return $out;
    }
}

$scriptDir = __DIR__;
$rootSchema = $scriptDir . '/native.xsd';
$outputMap = $scriptDir . '/schema_order_map.php';

if (!is_file($rootSchema)) {
    fwrite(STDERR, "ERROR: native.xsd not found at {$rootSchema}\n");
    exit(1);
}

$generator = new SchemaOrderMapGenerator($rootSchema);
$generator->run($outputMap);
