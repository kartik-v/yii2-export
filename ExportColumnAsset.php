<?php

/**
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2014
 * @package yii2-export
 * @version 1.2.0
 */

namespace kartik\export;

use kartik\widgets\AssetBundle;

/**
 * Asset bundle for ExportMenu Widget (for export columns selector)
 *
 * @author Kartik Visweswaran <kartikv2@gmail.com>
 * @since 1.0
 */
class ExportColumnAsset extends AssetBundle
{
    /**
     * @inheritdoc
     */
    public function init()
    {
        $this->setSourcePath(__DIR__ . '/assets');
        $this->setupAssets('js', ['js/kv-export-columns']);
        $this->setupAssets('css', ['css/kv-export-columns']);
        parent::init();
    }
}