<?php

use Phan\Issue;

return [

    'target_php_version' => null,
    'allow_missing_properties' => false,
    'null_casts_as_any_type' => true,
    'null_casts_as_array' => true,
    'array_casts_as_null' => true,
    'scalar_implicit_cast' => true,
    'scalar_array_key_cast' => true,
    'scalar_implicit_partial' => [],
    'strict_method_checking' => false,
    'strict_object_checking' => false,
    'strict_param_checking' => false,
    'strict_property_checking' => false,
    'strict_return_checking' => false,
    'ignore_undeclared_variables_in_global_scope' => false,
    'ignore_undeclared_functions_with_known_signatures' => true,
    'backward_compatibility_checks' => true, // this is valid for 5.6 -> 7 only
    'check_docblock_signature_return_type_match' => true,
    'phpdoc_type_mapping' => [],
    'dead_code_detection' => false,
    'unused_variable_detection' => false,
    'redundant_condition_detection' => false,
    'assume_real_types_for_internal_functions' => false,
    'quick_mode' => false,
    'globals_type_map' => [],
    'minimum_severity' => Issue::SEVERITY_CRITICAL,
    'suppress_issue_types' => [
        'PhanCompatibleVoidTypePHP70'
    ],
    'exclude_file_regex' => '@^vendor/.*/(tests?|Tests?)/@',
    'exclude_file_list' => [],
    'exclude_analysis_directory_list' => [
        'vendor/',
    ],
    'enable_include_path_checks' => false,
    'processes' => 1,
    'analyzed_file_extensions' => [
        'php',
    ],
    'autoload_internal_extension_signatures' => [],
    'plugins' => [],
    'directory_list' => [
        'src'
    ],
    'whitelist_issue_types' => [],
    'file_list' => [],
];
