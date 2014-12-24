<?php

/**
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2014
 * @package yii2-export
 * @version 1.1.0
 */

namespace kartik\export;

use kartik\widgets\AssetBundle;

/**
 * Asset bundle for GridView Widget (for exporting content)
 *
 * @author Kartik Visweswaran <kartikv2@gmail.com>
 * @since 1.0
 */
class ExportMenuAsset extends AssetBundle
{
    public function init()
    {
        $this->setSourcePath(__DIR__ . '/assets');
        $this->setupAssets('js', ['js/kv-export-data']);
        parent::init();
    }
}