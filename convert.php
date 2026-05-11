<?php

# Usage:
# php convert.php -x <xslx file> -f <files folder> [-v] [-l <default locale>] [-i]

# Hint: Use debugPrintXML($root) to write DOM to a file 'debug.xml'. $root should be a DOMDocument object.

// PHPExcel settings
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
define('EOL', (PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

require 'vendor/autoload.php';

class ConvertExcel2PKPNativeXML {

	// cli parsing
	private $opts;
	private $posArgs;
	private $fullFilesFolderPath;
	private $onlyValidate = false;
	private $ignoreMissingFiles = false;
	private $dummySubmissionFileName = '2017-1-1-1.pdf';
	private $dummySubmissionFilePath;
	private $dummySubmissionReplacementCount = 0;
	private $xlsxFileName = 'articleData.xlsx';
	private $filesFolderName = 'files';

	// defaults
	private $defaultUploader = 'admin';
	private $defaultAuthor = 'Editorial Board';
	private $defaultLocale = 'en';
	private $defaultUserGroupRef = [
		'en' => 'Author',
		'de' => 'Autor/in',
		'sv' => 'F&#xF6;rfattare'
	];
	private $primaryContactId;

	// table parsing
	private $issueKeys;
	private $sectionKeys;
	private $articleKeys;
	private $authorKeys;
	private $locales = [
		'en' => 'en',
		'fi' => 'fi',
		'sv' => 'sv',
		'de' => 'de',
		'ru' => 'ru',
		'fr' => 'fr',
		'no' => 'nb_NO',
		'da' => 'da',
		'es' => 'es',
	];	

	// xml generation
	private $articleElementOrder;
	private $publicationElementOrder;
	private $authorElementOrder;
	private $submissionFileElementOrder;
	private $issueElementOrder;
	private $issueIdentificationElementOrder;
	private $coverImageElementOrder;
	private $elementHasLocaleAttribute;
	
	// Constructor
	public function __construct($argv) {

		// pasre cli
		$rest_index = null;
		$shortOpts = "hvic:l:x:f:";
		$longOpts = ['defaultLocale:', 'validate','ignoreMissingFiles'];
		$this->opts = getopt($shortOpts, $longOpts, $rest_index);
		$this->posArgs = array_slice($argv, $rest_index);

		if (isset($this->opts['h']) || count($this->opts) == 0) {
			echo "Usage: php convert.php [-c <config.ini file>] [-x <xslx file>] [-f <files folder>] [-v] [-l <default locale>] [--ignoreMissingFiles|-i] [-h]", EOL;
			exit(0);
		}

		if (isset($this->opts['c'])) {
			$configFile = $this->opts['c'];
			if (!is_file($configFile)) {
				$this->logError("Config file does not exist");
				die();
			}
			// Parse the INI configuration file
			foreach (parse_ini_file($configFile, true) as $key => $value) {
				$this->{$key} = $value;
			}
		}

		if (!$this->validateInput()) {
			$this->logError("Data validation failed!");
		}

		// Get element order from build-time generated schema map
		$schemaOrderMapFile = __DIR__ . '/schema_order_map.php';
		if (!is_file($schemaOrderMapFile)) {
			$this->logError("schema order map not found. Run: php generate_schema_order_map.php");
			die();
		}

		$schemaOrderMap = require $schemaOrderMapFile;
		$this->articleElementOrder = $schemaOrderMap['articleElementOrder'];
		$this->publicationElementOrder = $schemaOrderMap['publicationElementOrder'];
		array_splice($this->publicationElementOrder, 1, 0, 'doi'); // allow 'doi' to come directly after 'id'
		$this->authorElementOrder = $schemaOrderMap['authorElementOrder'];
		$this->submissionFileElementOrder = $schemaOrderMap['submissionFileElementOrder'];
		$this->issueElementOrder = $schemaOrderMap['issueElementOrder'];
		$this->issueIdentificationElementOrder = $schemaOrderMap['issueIdentificationElementOrder'];
		$this->coverImageElementOrder = $schemaOrderMap['coverImageElementOrder'];
		$this->elementHasLocaleAttribute = $schemaOrderMap['elementsWithLocale'];

		// load data
		$this->logInfo("Reading Excel file");
		$objReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($this->xlsxFileName);
		$objReader->setReadDataOnly(false);
		$objPhpSpreadsheet = $objReader->load($this->xlsxFileName);
		$sheet = $objPhpSpreadsheet->setActiveSheetIndex(0);
		$articles = $this->createArray($sheet);

		/* 
		* Data validation   
		* -----------
		*/
		$msg = "Validating metadata from Excel sheet" . ($this->ignoreMissingFiles ? " (ignore missing files enabled)" : "");
		$this->logInfo($msg);

		$errors = $this->validateArticles($articles);
		if ($errors != "") {
			echo "\033[31m" . $errors . "\033[0m";
			die();
		}

		# Download galley files if fileName is empty or not provided and a gelleyDoi is available
		foreach ($articles as $index => &$article) {
			foreach ($article as $key => $value) {
				if (preg_match('/^fileName\d+$/', $key) && empty($value)) {
					$galleyDoiKey = str_replace('fileName', 'galleyDoi', $key);
					if (isset($article[$galleyDoiKey]) && !empty($article[$galleyDoiKey])) {
						$url = $article[$galleyDoiKey];
						$filename = basename($url) . '.pdf'; # PDF id just an assumption, should be improved
						$article[$key] = $filename;
						if (!is_file($this->fullFilesFolderPath.$filename)) {
							$fileContent = file_get_contents($url);
							if ($fileContent) {
								$article[$key] = $filename;
								file_put_contents($this->fullFilesFolderPath.$article[$key], $fileContent);
								$this->logInfo("Downloaded: ".$article[$key]." from $url");
							} else {
								$this->logError("Could not downlaod article galley from $url !");
							}
						}
					}
				}
			}
		}

		# If only validation is selected, exit
		if ($this->onlyValidate == 1) {
			$this->logInfo("Validation complete");
			die();
		}

		$this->process($articles);
	}

	function process($articles) {

		/* 
		* Prepare data for output
		* ----------------------------------------
		*/

		$this->logInfo("Preparing data for output");

		# Create issue and section identification
		$this->issueKeys = $this->getUniqueKeys($articles, 'issue');
		$this->sectionKeys = $this->getUniqueKeys($articles, 'section');
		$this->authorKeys = $this->getUniqueKeys($articles, 'author');
		$this->articleKeys = array_diff($this->getUniqueKeys($articles), $this->issueKeys, $this->sectionKeys, $this->authorKeys);

		$issueIdentifications = [];
		$issueData = [];
		$articles = array_values($articles); // reindex array (keys used as article id later)
		foreach ($articles as $id => $article) {

			// identify issue and sections for each article by hash generated from issue and section fields

			// get issue data
			$issueIdentification = array_intersect_key($article, array_flip($this->issueKeys));
			$articleIssueHash = hash("sha256", implode(
				", ",
				[
					array_key_exists('issueDatePublished',$issueIdentification)?$issueIdentification['issueDatePublished']:"",
					array_key_exists('issueVolume',$issueIdentification)?$issueIdentification['issueVolume']:"",
					array_key_exists('issueYear',$issueIdentification)?$issueIdentification['issueYear']:"",
					array_key_exists('issueTitle',$issueIdentification)?$issueIdentification['issueTitle']:""
				])
			);
			$issueIdentifications[$articleIssueHash] = $issueIdentification;
			foreach ($issueIdentification as $key => $value) {
				unset($article[$key]);
			}

			// get section data
			$sectionIdentification = array_intersect_key($article, array_flip($this->sectionKeys));
			// Sort the array alphabetically by the node value (required by native.xsd)
			ksort($sectionIdentification);
			foreach ($sectionIdentification as $key => $value) {
				unset($article[$key]);
			}

			// put all together			
			$issueData[$articleIssueHash]['issues'] = $issueIdentification;
			$issueData[$articleIssueHash]['sections'][$sectionIdentification['sectionAbbrev']] = $sectionIdentification;
			$issueData[$articleIssueHash]['articles'][$id] = $article;
			$issueData[$articleIssueHash]['articles'][$id]['sectionAbbrev'] = $sectionIdentification['sectionAbbrev'];
		}

		/* 
		* Create XML  
		* --------------------
		*/

		$this->logInfo("Starting XML output");

		$dom = new DOMDocument('1.0', 'UTF-8');
		$dom->formatOutput = true;

		[$issuesDOM, $pos] = $this->getOrCreateDOMElement($dom, 'issues', namespace: 'http://pkp.sfu.ca');
		$dom->appendChild($issuesDOM);

		// Create issue DOMs
		foreach ($issueData as $issueHash => $issueDataContent) {
			$issuesDOM = $this->processData($issuesDOM, $issueDataContent);
		}

		// reorder issue nodes
		foreach ($issuesDOM->childNodes as $issueDOM) {
			$issueDOM = $this->orderDOMNodes($issueDOM, $this->issueElementOrder);
		}

		$xpath = new DOMXPath($dom);
		$xpath->registerNamespace('xmlns', 'http://pkp.sfu.ca'); 
		$numberOfIssues = $xpath->query( "//xmlns:issue", $dom)->length;
		$numberOfSections = $xpath->query( "//section", $dom)->length;
		$numberOfArticles = $xpath->query( "//article", $dom)->length;
		$numberOfArticleGalleys = $xpath->query( "//xmlns:article_galley", $dom)->length;
		$numberOfEmbeds = $xpath->query( "//embed", $dom)->length;
		$this->logInfo("Added $numberOfIssues issues, $numberOfSections sections, $numberOfArticles articles, $numberOfArticleGalleys gelleys and $numberOfEmbeds embedded elements to the XML file.");

		$dom->save(filename: $this->xlsxFileName.'.xml');

		// validate XML
		// This currently will produce validation errors not reported by different other validation tools
		// Probably due to xsd version issues. Retry after PHP updated libxml
		// $this->validateXML($dom);
	}

	function validateXML($dom) {
		// Enable internal error handling
		libxml_use_internal_errors(true);

		// Validate the XML against the XSD
		$xsd = 'native.xsd';
		if ($dom->schemaValidate($xsd)) {
			echo "The XML document is valid.";
		} else {
			// Retrieve and display errors
			echo "The XML document is not valid. Errors:\n";
			foreach (libxml_get_errors() as $error) {
				echo "Error: " . $error->message;
				// echo "Line: " . $error->line . "\n";
				// echo "Column: " . $error->column . "\n";
			}
		}

		// Clear the error buffer
		libxml_clear_errors();
	}
	
	function validateInput() {
		// Check if the required parameter -x is set
		if (!isset($this->opts['x']) && !isset($this->xlsxFileName)) {
			$this->logError("Required parameter -x <xlsx filename> is missing.");
			exit(1);
		}
		if (isset($this->opts['x'])) $this->xlsxFileName = $this->opts['x'];

		// Check if the required parameter -f is set
		if (!isset($this->opts['f']) && !isset($this->filesFolderName)) {
			$this->logError("Required parameter -f <files folder name> is missing.");
			exit(1);
		}
		if (isset($this->opts['f'])) $this->filesFolderName = $this->opts['f'];

		// Check if the optional validate flag is set
		$this->onlyValidate = 0;
		if (isset($this->opts['v']) || isset($this->opts['validate'])) {
			$this->logInfo("Validation only mode is enabled.");
			$this->onlyValidate = 1;
		}

		// Check if ignoreMissingFiles mode is enabled.
		$ignoreFromOpts = isset($this->opts['i']) || isset($this->opts['ignoreMissingFiles']);
		$ignoreFromConfig = filter_var($this->ignoreMissingFiles, FILTER_VALIDATE_BOOLEAN);
		$this->ignoreMissingFiles = $ignoreFromOpts || $ignoreFromConfig;

		// Check if the defaultLocale optional parameter is set
		if (isset($this->opts['l'])) {
			$this->logInfo("Default locale is set with value: " . $this->opts['l']);
			$this->defaultLocale = $this->opts['l'];
		} elseif (isset($this->opts['defaultLocale'])) {
			$this->logInfo("Default locale is set with value: " . $this->opts['defaultLocale']);
			// The default locale. For alternative locales use language field. For additional locales use locale:fieldName.
			$this->defaultLocale = $this->opts['defaultLocale'];
		}

		/* 
		* Check that a file and a folder exists
		* ------------------------------------
		*/

		if (!is_file($this->xlsxFileName)) {
			$this->logError("Excel file does not exist");
			die();
		}

		// Location of full text files
		if (!str_starts_with($this->filesFolderName,'/')) {
			$this->filesFolderName = __DIR__ . "/" . $this->filesFolderName;
		}
		$this->fullFilesFolderPath = $this->filesFolderName . "/";

		if (!file_exists($this->fullFilesFolderPath)) {
			$this->logError("given folder does not exist");
			die();
		}

		if ($this->ignoreMissingFiles) {
			$this->dummySubmissionFilePath = $this->normalizePath(__DIR__ . '/exampleFiles/' . $this->dummySubmissionFileName);
			if (!is_file($this->dummySubmissionFilePath)) {
				$this->logError("ignoreMissingFiles is enabled but dummy file was not found: " . $this->dummySubmissionFilePath);
				die();
			}
			$this->logInfo("ignoreMissingFiles is enabled. Using dummy submission file: " . $this->dummySubmissionFilePath);
		}

		$this->logInfo("Basic input validation successful");

		return true;
	}

	function processData($dom, $data) {
		foreach ($data as $tagname => $content) {
			if (strlen($tagname) > 0) // to reject any blank lines in the excel sheet
			switch ($tagname) {
				case 'sections':
					[$issueDOM, $pos] = $this->getOrCreateDOMElement($dom->ownerDocument, 'issue');

					[$sectionsDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'sections');
					$issueDOM->appendChild($sectionsDOM);

					foreach ($data['sections'] as $sectionAbbrev => $sectionData) {
						[$sectionDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'section');
						$sectionDOM->setAttribute('abstract_word_count', 0);
						$sectionDOM = $this->processData($sectionDOM, $sectionData);	
						$sectionsDOM->appendChild($sectionDOM);
					}
					break;
				case 'issues':
					[$issueDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'issue', 'http://pkp.sfu.ca');
					$dom->appendChild($issueDOM);

					$issueData = $data[$tagname];
					$issueData = $this->stripColumnPrefix($issueData, 'issue');

					$issueIdentificationData = [];
					foreach ($this->issueIdentificationElementOrder as $field) {
						foreach (array_keys($issueData) as $key) {
							if (str_ends_with($key,$field)) {
								$issueIdentificationData[$key] = $issueData[$key];
								unset($issueData[$key]);
							}
						}
					}
					
					$issueData['id'] = [
						'type'=> 'internal',
						'id' => 0
					];
					$issueData['issue_identification'] = $issueIdentificationData;

					$issueData = $this->sortArrayElementsByKey($issueData, $this->issueElementOrder);

					$issueDOM = $this->processData($issueDOM, $issueData);

					$issueDOM->setAttribute('published', '1');
					$issueDOM->setAttribute('current', '0');

					$issueDOM = $this->orderDOMNodes($issueDOM, $this->issueElementOrder);
					break;
				case 'issue_identification':
					[$issuesIdentificationDOM, $pos] = $this->createDOMElement($dom->ownerDocument, $tagname);
					$dom->appendChild($issuesIdentificationDOM );

					$issuesIdentificationDOM = $this->processData($issuesIdentificationDOM, $content);
					break;
				case 'sectionTitle':
				case 'sectionAbbrev':
					// according to native.xsd abbrev and title need to be in (probably) alphabetic order (see ksort above)
					$xmlTagName = strtolower(str_replace('section', '', $tagname));
					$dom = $this->processData($dom, [$xmlTagName => $content]);
					$dom->setAttribute('ref', $data['sectionAbbrev']);
					$dom->setAttribute('seq', (int)(isset($data['sectionSeq']) ? $data['sectionSeq'] : "0"));
					break;
				case 'datePublished':
					[$issueDOM, $pos] = $this->getOrCreateDOMElement($dom->ownerDocument, 'issue');
					$element = $dom->ownerDocument->createElement('date_published', $content);
					$issueDOM->appendChild($element);
					$element = $dom->ownerDocument->createElement('last_modified', $content);
					$issueDOM->appendChild($element);
					break;
				case 'articles':
					[$articlesDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'articles', 'http://pkp.sfu.ca');
					[$issueDOM, $pos] = $this->getOrCreateDOMElement($dom->ownerDocument, 'issue');
					$issueDOM->appendChild($articlesDOM);

					foreach ($content as $articleId => $article) {

						# Article
							$this->logInfo("Adding article: " . $article['title']);
						[$articleDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'article');
						$articlesDOM->appendChild($articleDOM);

						$articleDOM->setAttribute('stage', 'production');
						$issueDatePublished = $dom->getElementsByTagName('date_published')[0]->textContent;
						$articleDOM->setAttribute('date_submitted', $issueDatePublished);
						$articleDOM->setAttribute('status', '3');
						$articleDOM->setAttribute('submission_progress', '0');
						$articleDOM = $this->processData($articleDOM, ['id' => [
								'type'=> 'internal',
								'id' => $articleId+1
							]]);

						// get file data 
						$fileKeys = $this->getUniqueKeys([$article], 'file');
						$fileData = [];
						foreach ($fileKeys as $key) {
							preg_match('/^(.*?)(\d+)$/', str_replace('file','',$key), $matches);
							$elementName = strtolower($matches[1]);
							$id = $matches[2];
							if (strlen($article[$key]) > 0) {
								if ($this->ignoreMissingFiles && $elementName === 'name') {
									$fileData[$id]['name'] = $this->dummySubmissionFileName;
									$this->dummySubmissionReplacementCount++;
								} elseif (
									(str_starts_with($article[$key], 'http://')) ||
									(str_starts_with($article[$key], 'https://'))
								) {
									// if file name is a url, extract the base file name
									$fileData[$id]['name'] = basename(parse_url($article[$key], PHP_URL_PATH)) . '.pdf'; # PDF id just an assumption, should be improved;
									$fileData[$id]['href'] = $article[$key];
								} else {
									$fileData[$id][$elementName] = $article[$key];
								}
							} else {
							}
							unset($article[$key]);
						}

						$articleDOM = $this->processData($articleDOM, [
							'submission_file' => $fileData
						]);
							
						$articleDOM = $this->processData($articleDOM, [
							'publication' => $article
						]);
					}

					if ($this->ignoreMissingFiles) {
						$this->logWarning("ignoreMissingFiles substituted " . $this->dummySubmissionReplacementCount . " submission files with dummy content");
					}
	
					break;
				case 'publication':
					[$publicationDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'publication');
					$dom->appendChild($publicationDOM);
					$publicationId = $pos;

					$dom->setAttribute('current_publication_id', $publicationId);
					
					# Check if language has an alternative default locale
					# If it does, use the locale in all fields
					$articleLocale = $this->defaultLocale;
					if (!empty($content['language'])) {
						$articleLocale = $this->locales[trim($content['language'])];
					}
					unset($content['language']);

					$publicationDOM->setAttribute('version', "1");
					$publicationDOM->setAttribute('status', "3");
					$publicationDOM->setAttribute('access_status', "0");
					$publicationDOM->setAttribute('url_path', "");

					[$element, $pos] = $this->getOrCreateDOMElement($dom->ownerDocument,'issue');
					$datePublishedList = $element->getElementsByTagName('date_published');
					if ($datePublishedList->length > 0) {
						$publicationDOM->setAttribute('date_published', $datePublishedList[0]->textContent);
					}
					
					$publicationDOM->setAttribute('section_ref', $content['sectionAbbrev']);
					unset($content['sectionAbbrev']);
					if (isset($content['articleSeq'])) {
						$publicationDOM->setAttribute('seq', (int)$content['articleSeq']);
						unset($content['articleSeq']);
					}

					$publicationDOM = $this->processData($publicationDOM,  [
						'id' => [
							'type'=> 'internal',
							'id' => $publicationId
						]]);

					$content = $this->stripColumnPrefix($content, 'article');

					//  get the author data (we need to process it later)
					$authorKeys = $this->getUniqueKeys([$content], 'author');
					$content['authors'] = [];
					foreach ($authorKeys as $key) {
						preg_match('/^(.*?)(\d+)$/', str_replace('author','',$key), $matches);
						$elementName = strtolower($matches[1]);
						$id = $matches[2];
						if (strlen($content[$key]) > 0) {
							$content['authors'][$id][$elementName] = $content[$key];
						}
						unset($content[$key]);
					}
				
				// Extract localized author affiliations (e.g., en:authorAffiliation1)
				$localizedAuthorAffiliations = [];
				foreach ($content as $key => $value) {
					if (preg_match('/^[a-z]{2}:author(Affiliation|RorAffiliation)(\d+)$/', $key, $matches)) {
						$affiliationType = $matches[1]; // 'Affiliation' or 'RorAffiliation'
						$authorId = $matches[2];
						[$locale, $fieldName] = $this->splitLocaleTagName($key);
						if (strlen($value) > 0) {
							if ($affiliationType === 'Affiliation') {
								if (!isset($localizedAuthorAffiliations[$authorId]['affiliation'])) {
									$localizedAuthorAffiliations[$authorId]['affiliation'] = [];
								}
								$localizedAuthorAffiliations[$authorId]['affiliation'][$locale] = $value;
							}
						}
						unset($content[$key]);
					}
				}
				
				foreach ($content['authors'] as $id => $authorData) {
					// Merge localized affiliations if present
					if (isset($localizedAuthorAffiliations[$id]['affiliation'])) {
						$authorData['affiliation'] = $localizedAuthorAffiliations[$id]['affiliation'];
					} else {
						// Localized author affiliation columns can be extracted as locale-prefixed keys (e.g., en:affiliation).
						// Fold them back into a single compound affiliation structure.
						$localizedAffiliationValues = [];
						foreach ($authorData as $field => $fieldValue) {
							if (preg_match('/^([a-z]{2}):affiliation$/', $field, $matches)) {
								$localePrefix = strtolower($matches[1]);
								$locale = $this->locales[$localePrefix] ?? $this->defaultLocale;
								if (strlen((string)$fieldValue) > 0) {
									$localizedAffiliationValues[$locale] = $fieldValue;
								}
								unset($authorData[$field]);
							}
						}

						if (!empty($localizedAffiliationValues)) {
							$authorData['affiliation'] = $localizedAffiliationValues;
						}
					}
					
					// create required fields if not provided
					$missingKeys = array_diff(
						['givenname','familyname','affiliation','country','email'],
						array_keys($authorData)
					);
					// set missing values
					foreach ($missingKeys as $key) {
						if ($key == 'givenname') {
							$authorData[$key] = $this->defaultAuthor;
						} else {
							$authorData[$key] = "";
						}
					}
					// sort elements according to required field order
					$content['authors'][$id] = $this->sortArrayElementsByKey($authorData, $this->authorElementOrder);
				}

					// get galley data 
					$galleyKeys = $this->getUniqueKeys([$content], 'galley');
					$galleyData = [];
					foreach ($galleyKeys as $key) {
						preg_match('/^(.*?)(\d+)$/', str_replace('galley','',$key), $matches);
						$elementName = strtolower($matches[1]);
						$id = $matches[2];
						if (strlen($content[$key]) > 0) {
							$galleyData[$id][$elementName] = $content[$key];
						}
						unset($content[$key]);
					}
					$content['article_galley'] = $galleyData;

					$content = $this->sortArrayElementsByKey($content, $this->publicationElementOrder);

					// process data
					$publicationDOM = $this->processData($publicationDOM, $content);

					$publicationDOM->setAttribute('primary_contact_id', $this->primaryContactId);
					break;
				case 'authors':
					[$authorsDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'authors', 'http://pkp.sfu.ca');
					$dom->appendChild($authorsDOM);
					$this->primaryContactId = null;

					$requestedPrimaryContactId = isset($data['primaryContactId']) ? (string)$data['primaryContactId'] : null;
					unset($data['primaryContactId']);

					foreach ($content as $authorId => $author) {
							
						[$authorDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'author');
						$authorsDOM->appendChild($authorDOM);
						$authorDomId = $pos + 1;
				
						$authorDOM->setAttribute('include_in_browse', 'true');
						$authorDOM->setAttribute('user_group_ref', $this->defaultUserGroupRef[$this->defaultLocale]);
						$authorDOM->setAttribute('seq', (int)$authorId);
						$authorDOM->setAttribute('id', $authorDomId);

						if (!isset($this->primaryContactId)) {
							// Default to the first author when no explicit primary contact is provided.
							$this->primaryContactId = $authorDomId;
						}

						if ($requestedPrimaryContactId !== null && $requestedPrimaryContactId === (string)$authorId) {
							$this->primaryContactId = $authorDomId;
						}

						$authorDOM = $this->processData($authorDOM, $author);
					}
					break;
				case 'affiliation':
					// Check if this is a compound affiliation structure (array of locales) or simple string
					if (is_array($content)) {
						// Compound affiliation with multiple locales
						$this->processCompoundAffiliation($dom, $content);
					} else {
						// Simple affiliation (backward compatibility)
						$affiliations = array_filter(array_map('trim', explode(';', (string)$content)), fn($value) => $value !== '');
						foreach ($affiliations as $affiliationName) {
							[$affiliationDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'affiliation');
							$dom->appendChild($affiliationDOM);
							$affiliationDOM = $this->processData($affiliationDOM, ['name' => $affiliationName]);
						}
					}
					break;
				case 'roraffiliation':
					$rorAffiliations = array_filter(array_map('trim', explode(';', (string)$content)), fn($value) => $value !== '');
					foreach ($rorAffiliations as $rorAffiliation) {
						$parts = array_map('trim', explode('|', $rorAffiliation, 2));
						$rorValue = $parts[0] ?? '';
						$nameValue = $parts[1] ?? '';

						if ($rorValue === '' || $nameValue === '') {
							continue;
						}

						[$rorAffiliationDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'rorAffiliation');
						$dom->appendChild($rorAffiliationDOM);
						$rorAffiliationDOM = $this->processData($rorAffiliationDOM, [
							'ror' => $rorValue,
							'name' => $nameValue
						]);
					}
					break;
				case 'article_galley':
					foreach ($content as $id => $galleyData) {
						[$articleGalleysDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'article_galley', 'http://pkp.sfu.ca');
						$dom->appendChild($articleGalleysDOM);
						
						$articleGalleysDOM->setAttribute('locale', $this->locales[$galleyData['locale']]);
						$articleGalleysDOM->setAttribute('approved', "false");

						if (array_key_exists('doi', $galleyData)) {
							$articleGalleysDOM = $this->processData($articleGalleysDOM, ['doi' => $galleyData['doi']]);
						}
						// When importing OJS seems to ignore the <article_galley> locale attribute but set the galley locale on the basis of the locale of the <name> tag
						$articleGalleysDOM = $this->processData($articleGalleysDOM, [$galleyData['locale'].':name' => $galleyData['label']]);
						$articleGalleysDOM = $this->processData($articleGalleysDOM, ['seq' => (int)($id-1)]);

						[$fileRef, $pos] = $this->createDOMElement($dom->ownerDocument, 'submission_file_ref'); # Todo: This element should be replace in case of remote galleys
						$articleGalleysDOM->appendChild($fileRef);

						$fileRef->setAttribute('id', $pos+1);
					}
					break;
				case 'id':
					switch ($content['type']) {
						case 'internal':
							$id = $dom->ownerDocument->createElement('id', $content['id']);
							$id->setAttribute('type', 'internal');
							$id->setAttribute('advice', 'ignore');
							break;
					}
					$dom->appendChild($id);
					break;
				case 'doi':
					$id = $dom->ownerDocument->createElement('id', $content);
					$dom->appendChild($id);

					$id->setAttribute('type', 'doi');
					$id->setAttribute('advice', 'update');
					break;
				case str_ends_with($tagname, 'keywords'):
				case str_ends_with($tagname, 'keyword'):
				case str_ends_with($tagname, 'disciplines'):
				case str_ends_with($tagname, 'discipline'):
				case str_ends_with($tagname, 'subjects'):
				case str_ends_with($tagname, 'subject'):
				case str_ends_with($tagname, 'references'):
				case str_ends_with($tagname, 'citations'):
				case str_ends_with($tagname, 'citation'):
					if (str_ends_with($tagname, 'references')) {
						$tagname = str_replace('references', 'citations', $tagname); // citations is the correct tag name according to native.xsd
					} elseif (str_ends_with($tagname, 'discipline')) {
						$tagname = preg_replace('/discipline$/', 'disciplines', $tagname);
					} elseif (str_ends_with($tagname, 'subject')) {
						$tagname = preg_replace('/subject$/', 'subjects', $tagname);
					} elseif (str_ends_with($tagname, 'keyword')) {
						$tagname = preg_replace('/keyword$/', 'keywords', $tagname);
					} elseif (str_ends_with($tagname, 'citation')) {
						$tagname = preg_replace('/citation$/', 'citations', $tagname);
					}
					if (strlen((string)$content) > 0) {
						[$locale, $xmlTagName] = $this->splitLocaleTagName($tagname);
						[$elementsDOM, $pos] = $this->createDOMElement($dom->ownerDocument, $xmlTagName);
						$dom->appendChild($elementsDOM);
	
						$elementsDOM->setAttribute('locale', $locale);
						
						// Schema-specific child tags:
						// - keywords/disciplines/subjects require <keyword><name>...</name></keyword>
						// - citations require <citation>...</citation>
						foreach (array_filter(array_map('trim', explode(';', (string)$content)), fn($value) => $value !== '') as $element) {
							if ($xmlTagName === 'citations') {
								[$elementDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'citation');
								$elementDOM->appendChild($dom->ownerDocument->createTextNode($element));
							} else {
								[$elementDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'keyword');
								$nameDOM = $dom->ownerDocument->createElement('name', $element);
								$elementDOM->appendChild($nameDOM);
							}
							$elementsDOM->appendChild($elementDOM);
						}
					}
					break;
				case 'submission_file':
					foreach ($content as $id => $submissionFileData) {
						[$subFileDOM, $pos] = $this->createDOMElement($dom->ownerDocument, $tagname, 'http://pkp.sfu.ca');
						$dom->appendChild($subFileDOM);

						$subFileDOM->setAttribute('stage', 'proof');
						$subFileDOM->setAttribute('id', $pos+1);
						$subFileDOM->setAttribute('file_id', $pos+1);
						$subFileDOM->setAttribute('uploader', $this->defaultUploader);
						$subFileDOM->setAttribute('genre', $submissionFileData['genre']);
						unset($submissionFileData['genre']);

						$submissionFileData = $this->sortArrayElementsByKey($submissionFileData, $this->submissionFileElementOrder);
	
						# If href is provided, it will be handled below
						$submissionFileDataNoHref = $submissionFileData;
						unset($submissionFileDataNoHref['href']);
						$subFileDOM = $this->processData($subFileDOM, $submissionFileDataNoHref);
						
						# create file element
						$filePath = $this->normalizePath($this->fullFilesFolderPath . $submissionFileData['name']);
						if (is_file($filePath) || array_key_exists('href', $submissionFileData)) {
							$this->logInfo("Adding file " . $filePath);
							$file = $dom->ownerDocument->createElement('file');
							$subFileDOM->appendChild($file);

							$file->setAttribute('id', $pos+1);				
							$file->setAttribute('extension', pathinfo($submissionFileData['name'], PATHINFO_EXTENSION));

							if (array_key_exists('href', $submissionFileData)) {
								# link to external file
								$file = $this->processData($file, ['href' => $submissionFileData['href']]);
							} else {
								# embed file content as base64
								$size = filesize($filePath);
								$file->setAttribute('filesize', $size);
								$embed = $dom->ownerDocument->createElement('embed', base64_encode(file_get_contents($filePath)));
								$embed->setAttribute('encoding','base64');
								$file->appendChild($embed);
							}
						} else {
							$this->logWarning("file " . $filePath . " not found !");
						}
					}
					break;
				case str_ends_with($tagname, 'coverImage'):
				case str_ends_with($tagname, 'coverImageAltText'):
					if (strlen($content) > 0) {

						[$locale, $xmlTagName] = $this->splitLocaleTagName($tagname);
						[$coverDOM, $pos] = $this->getOrCreateDOMElement($dom, 'cover', $locale);
						if ($coverDOM->childElementCount == 0) {
							[$coversDOM, $pos] = $this->getOrCreateDOMElement($dom, 'covers', namespace: 'http://pkp.sfu.ca');
							$coversDOM->appendChild($coverDOM);
							$dom->appendChild($coversDOM);
							$coverDOM->setAttribute('locale', $locale);
						}

						if ($xmlTagName == 'coverImageAltText') {
							$node = $coverDOM->getElementsByTagName('cover_image_alt_text')[0];
							if ($node) {
								$coverDOM->removeChild($node);
							}
							$coverDOM = $this->processData($coverDOM, ['cover_image_alt_text' => $content]);
						} else {
							$coverDOM = $this->processData($coverDOM, ['cover_image' => $content]);
							$filePath = $this->fullFilesFolderPath . $content;
							if (is_file($filePath)) {
								$embed = $dom->ownerDocument->createElement('embed', base64_encode(file_get_contents($filePath)));
								$embed->setAttribute('encoding','base64');
								$coverDOM->appendChild($embed);
							}
						}

						//reorder nodes and set default alt text
						$childNodes = ['cover_image_alt_text' => $dom->ownerDocument->createElement('cover_image_alt_text',"")];
						foreach ($coverDOM->childNodes as $child) {
							$childNodes[$child->tagName] = $child;
						}
						$childNodes = $this->sortArrayElementsByKey($childNodes, $this->coverImageElementOrder);
						foreach ($childNodes as $child) {
							$coverDOM->appendChild($child);
						}
					}
					break;
				case 'href':
					$element = $dom->ownerDocument->createElement($tagname);
					$element->setAttribute('src', $content);
					$dom->appendChild($element);
					break;
				default:
					// here we handle all text nodes
					if (strlen($tagname) > 0) {
						switch ($tagname) {
							case 'primaryContactId':
							case 'sectionSeq':
								// fields that hold attributes don't create a tag
								break;
							case in_array($tagname, $this->elementHasLocaleAttribute):
							case (strpos($tagname, ':') === 2):
								// elements with locale attribute
								[$locale, $tagname] = $this->splitLocaleTagName($tagname);

								$element = $dom->ownerDocument->createElement($tagname);
								$element->appendChild($dom->ownerDocument->createTextNode($content));
								if ($locale) {
									$element->setAttribute('locale', $locale);
								}
								
								$dom->appendChild($element);
								break;
							case 'copyrightYear':
								if (strlen($content) == 0) {
									break;
								}
							default:
								// elements without locale attribute
								$element = $dom->ownerDocument->createElement($tagname);
								$element->appendChild($dom->ownerDocument->createTextNode($content));
								$dom->appendChild($element);
								break;
						}
					}
					break;
			}
		}
		return $dom;
	}

	/* 
	* Helpers 
	* -----------
	*/

	private function logInfo($message) {
		echo date('H:i:s') . " " . $message . EOL;
	}

	private function logWarning($message) {
		echo "\033[1;33m" . date('H:i:s') . " " . $message . "\033[0m" . EOL;
	}

	private function logError($message) {
		echo "\033[31m" . date('H:i:s') . " " . $message . "\033[0m" . EOL;
	}

	// sort elements according to required field order
	function sortArrayElementsByKey(array $dataArray, array $fieldOrder) {

		$orderedArray = [];
		foreach ($fieldOrder as $xmlTagName) {
			if (isset($dataArray[$xmlTagName])) {
				$orderedArray[$xmlTagName] =  $dataArray[$xmlTagName];
				unset($dataArray[$xmlTagName]);
			}
			foreach (array_flip($this->locales) as $locale) {
				$localeKey = $locale.':'.$xmlTagName;
				if (array_key_exists($localeKey, $dataArray)) {
					$orderedArray[$localeKey] = $dataArray[$localeKey];
					unset($dataArray[$localeKey]);
				}
			}
		}
		$orderedArray = array_merge($orderedArray, $dataArray);
		return $orderedArray;
	}

	// extract locale value from tag name
	function splitLocaleTagName($tagname, $locale = NULL) {
		// Is there a valid locale specified?
		if (strpos($tagname, ":") !== false) {
			$locale = $this->locales[explode(':',$tagname)[0]];
			if (!$locale) {
				$locale = $this->defaultLocale;
			} else {
				$tagname = explode(':',$tagname)[1];
			}
			return [$locale, $tagname];
		} else {
			$locale = $this->defaultLocale;
		}
		return [$locale, $tagname];
	}

	function createDOMElement($root, $tagname, $namespace = NULL) {
		$this->logInfo("Creating element " . $tagname);
		if ($namespace) {
			$targetDOM = $root->createElementNS($namespace, $tagname);
			$targetDOM->setAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
			$targetDOM->setAttribute('xsi:schemaLocation', 'http://pkp.sfu.ca native.xsd');
		} else {
			$targetDOM = $root->createElement($tagname);
		}
		
		return [$targetDOM, $root->getElementsByTagname($tagname)->length];
	}

	// Try to get the last DOM element with the given tag name. If none exists create a new one
	function getOrCreateDOMElement($dom, $tagname, $locale = NULL, $namespace = NULL) {
		// try to get the requested element
		if (get_class($dom) !== "DOMDocument") {
			$root = $dom->ownerDocument;
		} else {
			$root = $dom;
		}
		$targetDOM = $dom->getElementsByTagName($tagname);
		if (($targetDOM->length > 0) && $locale) {
			$xpath = new DOMXPath($root);
			$targetDOM = $xpath->query(
				expression: "//".$tagname."[@locale='$locale']",
				contextNode: $dom
			); 
		}
		
		// create element if not found
		if (!isset($targetDOM) || $targetDOM->length == 0) {
			[$targetDOM, $pos] = $this->createDOMElement($root, $tagname, $namespace);
			if (get_class($dom) !== "DOMDocument") {
				$dom->lastChild->appendChild($targetDOM);
			} else {
				$dom->appendChild($targetDOM);
			}
			$elementPosition = $pos;
		} else {
			$elementPosition = $targetDOM->length;
			$targetDOM = $targetDOM->item($targetDOM->length - 1);
		}

		return [$targetDOM, $elementPosition];
	}

	# Function for creating an array using the first row as keys
	function createArray($sheet)
	{
		$highestrow = $sheet->getHighestRow();
		$highestcolumn = $sheet->getHighestColumn();
		$columncount = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestcolumn);
		$headerRow = $sheet->rangeToArray('A1:' . $highestcolumn . "1");
		$header = $headerRow[0];
		array_unshift($header, "");
		unset($header[0]);
		$array = array();
		for ($row = 2; $row <= $highestrow; $row++) {
			$a = array();
			for ($column = 1; $column <= $columncount; $column++) {
				$cellCoordinate = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($column) . $row;
				$cell = $sheet->getCell($cellCoordinate);
				if (strpos($header[$column], "abstract") !== false) {
					if ($cell->getValue() instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
						$value = $cell->getValue();
						$elements = $value->getRichTextElements();
						$cellData = "";
						foreach ($elements as $element) {
							if ($element instanceof \PhpOffice\PhpSpreadsheet\RichText\Run) {
								if ($element->getFont()->getBold()) {
									$cellData .= '<b>';
								} elseif ($element->getFont()->getSubScript()) {
									$cellData .= '<sub>';
								} elseif ($element->getFont()->getSuperScript()) {
									$cellData .= '<sup>';
								} elseif ($element->getFont()->getItalic()) {
									$cellData .= '<em>';
								}
							}
							// Convert UTF8 data to PCDATA
							$cellText = $element->getText();
							$cellData .= htmlspecialchars($cellText);
							if ($element instanceof \PhpOffice\PhpSpreadsheet\RichText\Run) {
								if ($element->getFont()->getBold()) {
									$cellData .= '</b>';
								} elseif ($element->getFont()->getSubScript()) {
									$cellData .= '</sub>';
								} elseif ($element->getFont()->getSuperScript()) {
									$cellData .= '</sup>';
								} elseif ($element->getFont()->getItalic()) {
									$cellData .= '</em>';
								}
							}
						}
						$a[$header[$column]] = $cellData;
					} else {
						$a[$header[$column]] = $cell->getFormattedValue();
					}
				} else {
					$key = $header[$column];
					$a[$key] = $cell->getFormattedValue();
				}
			}
			$array[$row] = $a;
		}

		return $array;
	}

	# Function for data validation
	function validateArticles($articles)
	{
		$errors = "";
		$articleRow = 0;

		foreach ($articles as $article) {

			$articleRow++;

			if (empty($article['issueYear'])) {
				$errors .= date('H:i:s') . " ERROR: Issue year missing for article " . $articleRow . EOL;
			}

			if (empty($article['issueDatePublished'])) {
				$errors .= date('H:i:s') . " ERROR: Issue publication date missing for article " . $articleRow . EOL;
			}

			if (empty($article['title'])) {
				$errors .= date('H:i:s') . " ERROR: article title missing for the given default locale for article " . $articleRow . EOL;
			}

			if (empty($article['sectionTitle'])) {
				$errors .= date('H:i:s') . " ERROR: section title missing for the given default locale for article " . $articleRow . EOL;
			}

			if (empty($article['sectionAbbrev'])) {
				$errors .= date('H:i:s') . " ERROR: section abbreviation missing for the given default locale for article " . $articleRow . EOL;
			}

			$articleLanguage = strtolower(trim($article['language'] ?? ''));
			$localePrefix = '';
			foreach ($this->locales as $lang => $localeCode) {
				if (strtolower($lang) === $articleLanguage || strtolower($localeCode) === $articleLanguage) {
					$localePrefix = $lang;
					break;
				}
			}

			$defaultLocalePrefix = '';
			foreach ($this->locales as $lang => $localeCode) {
				if (strtolower($localeCode) === strtolower($this->defaultLocale) || strtolower($lang) === strtolower($this->defaultLocale)) {
					$defaultLocalePrefix = $lang;
					break;
				}
			}

			$requiredLocalePrefix = $localePrefix !== '' ? $localePrefix : $defaultLocalePrefix;

			// prepare all author-related fields
			$authorKeys = $this->getUniqueKeys([$article], 'author');
			foreach ($authorKeys as $authorKey) {
				$authorIndex = (int) filter_var($authorKey, FILTER_SANITIZE_NUMBER_INT);

				$authorGivenName = trim((string)($article['authorGivenname' . $authorIndex] ?? ''));
				$authorFamilyName = trim((string)($article['authorFamilyname' . $authorIndex] ?? ''));
				$authorEmail = trim((string)($article['authorEmail' . $authorIndex] ?? ''));

				if ($authorGivenName === '' && $authorFamilyName === '' && $authorEmail === '') {
					break;
				}

				$hasGenericAffiliation = trim((string)($article['authorAffiliation' . $authorIndex] ?? '')) !== ''
					|| trim((string)($article['authorRorAffiliation' . $authorIndex] ?? '')) !== '';

				$hasAnyAffiliationData = $hasGenericAffiliation;
				if (!$hasAnyAffiliationData) {
					foreach ($article as $fieldName => $fieldValue) {
						if (trim((string)$fieldValue) === '') {
							continue;
						}

						if (preg_match('/^[a-z]{2}:(authorAffiliation|authorRorAffiliation)' . $authorIndex . '$/i', $fieldName)) {
							$hasAnyAffiliationData = true;
							break;
						}
					}
				}

				$hasLocalizedAffiliation = false;
				if ($requiredLocalePrefix !== '') {
					$localizedAffiliationKey = $requiredLocalePrefix . ':authorAffiliation' . $authorIndex;
					$localizedRorAffiliationKey = $requiredLocalePrefix . ':authorRorAffiliation' . $authorIndex;
					$hasLocalizedAffiliation = trim((string)($article[$localizedAffiliationKey] ?? '')) !== ''
						|| trim((string)($article[$localizedRorAffiliationKey] ?? '')) !== '';
				}

				$allowGenericForSubmissionLocale = $requiredLocalePrefix === '' || $requiredLocalePrefix === $defaultLocalePrefix;
				$hasAffiliationForSubmissionLocale = $allowGenericForSubmissionLocale
					? ($hasGenericAffiliation || $hasLocalizedAffiliation)
					: $hasLocalizedAffiliation;

				if ($hasAnyAffiliationData && !$hasAffiliationForSubmissionLocale) {
					$languageLabel = $requiredLocalePrefix !== '' ? $requiredLocalePrefix : ($articleLanguage !== '' ? $articleLanguage : $this->defaultLocale);
					$errors .= date('H:i:s') . " ERROR: author " . $authorIndex . " affiliation is provided but missing for submission language " . $languageLabel . " for article " . $articleRow . EOL;
				}
			}

			// get all array keys starting with "fileName" to check if the referenced files exist and if galley label and locale are provided
			$fileKeys = $this->getUniqueKeys([$article], 'fileName');
			foreach ($fileKeys as $fileKey) {
				$i = (int) filter_var($fileKey, FILTER_SANITIZE_NUMBER_INT);
				$fileName = $article['fileName' . $i] ?? '';

				if ($fileName === '') {
					break;
				}

				if (!preg_match("@^https?://@", $fileName)) {
					$fileCheck = $this->fullFilesFolderPath . $fileName;
					if (!is_file($fileCheck) && !$this->ignoreMissingFiles) {
						$errors .= date('H:i:s') . " ERROR: file " . $i . " missing " . $fileCheck . EOL;
					}
				}

				if (empty($article['galleyLabel' . $i])) {
					$errors .= date('H:i:s') . " ERROR: galleyLabel " . $i . " missing for article " . $articleRow . EOL;
				}
				if (empty($article['galleyLocale' . $i])) {
					$errors .= date('H:i:s') . " ERROR: galleyLocale " . $i . "  missingfor article " . $articleRow . EOL;
				}
			}
		}

		return $errors;
	}

	// get unique column keys starting with <name>
	function getUniqueKeys($articles, $name = NULL)
	{
		$uniqueKeys = array_unique(array_keys(array_merge(...$articles)));
		if (!$name) {
			return $uniqueKeys;
		}
		return array_filter($uniqueKeys, function ($key) use ($name) {
			return ((strpos($key, $name) === 0) || (strpos($key, $name) === 3)); // does the name occur at the beginning or after a locale code?
		});
	}

	// Function to find the specific element and get its child elements
	function getChildElementsOrder($node, $elementName) {
		$order = [];

		// Check if the node is the target element
		if ($node->nodeType === XML_ELEMENT_NODE && $node->nodeName === $elementName) {
			if ($node->hasChildNodes()) {
				foreach ($node->childNodes as $child) {
					if ($child->nodeType === XML_ELEMENT_NODE) {
						$order[] = $child->nodeName; // Add the element name to the order
					}
				}
			}
		}

		// Recursively search for the target element in child nodes
		if ($node->hasChildNodes()) {
			foreach ($node->childNodes as $child) {
				$order = array_merge($order, $this->getChildElementsOrder($child, $elementName));
			}
		}

		return array_unique($order);
	}

	// return an array with element names that have the "locale" attribute
	function hasLocaleAttribute($dom) {

		$elementsWithLocale = [];

		$elements = $dom->getElementsByTagName('*');
		foreach ($elements as $element) {
			if ($element->hasAttribute('locale')) {
				$elementsWithLocale[] = $element->nodeName;
			}
		}
		return $elementsWithLocale;
	}

	// strip 'article' prefix from key names
	function stripColumnPrefix(array $data, string $prefix) {
		foreach ($data as $key => $value) {
			// Check if the key starts with 'article'
			if ((strpos($key, $prefix) === 0) || (strpos($key, $prefix) === 3)) {
				// Remove 'article' from the key
				if (strpos($key, $prefix) === 3) {
					// there is a locale descriptor we need to consider
					$newKey = str_replace($prefix, '', $key);
					$keyParts = explode(':', $newKey);
					$keyParts[1]= lcfirst($keyParts[1]);
					$key = implode(':', $keyParts);
					$newKey = str_replace($prefix, '', $key);
				} else {
					$newKey = lcfirst(str_replace($prefix, '', $key));
				}
				$newArray[$newKey] = $value; // Assign the value to the new key
			} else {
				$newArray[$key] = $value; // Keep the original key-value pair
			}
		}
		return $newArray;
	}

	function orderDOMNodes($dom, $order) {
		$childNodes = [];
		foreach ($dom->childNodes as $key => $child) {
			$childNodes[$child->tagName][$key] = $child;
		}
		$childNodes = $this->sortArrayElementsByKey($childNodes, $order);
		foreach ($childNodes as $tagname) {
			foreach ($tagname as $child) {
				$dom->appendChild($child);
			}
		}
		return $dom;
	}

	function normalizePath($path) {
		// Replace backslashes with forward slashes
		$normalizedPath = str_replace('\\', '/', $path);
		
		// Remove any redundant slashes
		$normalizedPath = preg_replace('~/+~', '/', $normalizedPath);
		
		return $normalizedPath;
	}

	/**
	 * Process compound affiliation structures with multiple locales
	 * Creates a single <affiliation> element with multiple <name> children, each with a locale attribute
	 * 
	 * @param DOMElement $dom The parent DOM element to append affiliation to
	 * @param array $content Associative array mapping locale codes to affiliation strings
	 *                       Example: ['en' => 'University A; University B', 'de' => 'Universität A']
	 */
	function processCompoundAffiliation($dom, $content) {
		// Group all affiliation names by locale
		$affiliationsByLocale = [];
		
		foreach ($content as $locale => $affiliationText) {
			if (is_string($affiliationText) && strlen($affiliationText) > 0) {
				// Split by semicolon to support multiple affiliations per locale
				$affiliations = array_filter(
					array_map('trim', explode(';', $affiliationText)),
					fn($value) => $value !== ''
				);
				if (!empty($affiliations)) {
					$affiliationsByLocale[$locale] = $affiliations;
				}
			}
		}
		
		// Create a single affiliation element with multiple name children
		if (!empty($affiliationsByLocale)) {
			[$affiliationDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'affiliation');
			$dom->appendChild($affiliationDOM);
			
			foreach ($affiliationsByLocale as $locale => $affiliationNames) {
				foreach ($affiliationNames as $affiliationName) {
					$nameElement = $dom->ownerDocument->createElement('name');
					$nameElement->appendChild($dom->ownerDocument->createTextNode($affiliationName));
					$nameElement->setAttribute('locale', $locale);
					$affiliationDOM->appendChild($nameElement);
				}
			}
		}
	}

}

$app = new ConvertExcel2PKPNativeXML($argv);

function debugPrintXML($root) {
	$root->save('debug.xml');
}