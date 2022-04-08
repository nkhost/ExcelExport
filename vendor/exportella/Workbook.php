<?php

namespace Exportella;

use \Exception;

/**
 * Книга Xlsx документа
 */
class Workbook
{
  /**
   * Путь к шаблону
   * @var string
   */
  protected $templatePath;
  
  /**
   * Путь к временной папке
   * @var string
   */
  protected $tmpDir;
  
  /**
   * Путь к папке - распакованному xlsx документу
   * @var string
   */
  protected $unpackedDir;
  
  /**
   * Путь к файлу workbook.xml
   * @var string
   */
  protected $workbookPath;
  
  /**
   * Объект для работы с sharedStrings.xml
   * @var SharedStrings
   */
  protected $sharedStrings;
  
  /**
   * @param string $templatePath Путь к шаблону
   */
  public function __construct(string $templatePath, ?string $tmpDir = null)
  {
    $this->templatePath = $templatePath;
    $this->tmpDir = $tmpDir ?? sys_get_temp_dir();
  }
  
  /**
   * Распаковать xlsx документ Во временную папку
   *
   * @return void
   * @throws Exception
   */
  public function extract()
  {
    $zip = new PclZip($this->templatePath);
    $this->unpackedDir = $this->tmpDir . '/' . uniqid('exportella_');
    if (empty($zip->extract(PclZip::PCLZIP_OPT_REPLACE_NEWER, PclZip::PCLZIP_OPT_PATH, $this->unpackedDir))) {
      throw new Exception('Unable to extract xlsx file');
    }
    $this->workbookPath = $this->unpackedDir . '/xl/workbook.xml';
  }
  
  /**
   * Получить объект листа по номеру
   *
   * @param int $number Номер листа
   * @return Worksheet
   */
  public function getWorksheet(int $number): Worksheet
  {
    return new Worksheet($this->unpackedDir . '/xl/worksheets/sheet' . $number . '.xml', $this);
  }
  
  /**
   * Получить объект работы с SharedStrings.xml
   *
   * @return SharedStrings
   * @throws Exception
   */
  public function getSharedStrings(): SharedStrings
  {
    if (!$this->sharedStrings) {
      $this->sharedStrings = new SharedStrings($this->unpackedDir . '/xl/sharedStrings.xml', $this);
      $this->sharedStrings->load();
    }
    return $this->sharedStrings;
  }
  
  /**
   * Создать xlsx документ
   *
   * @param string $destination Путь назначения файла
   * @return void
   * @throws Exception
   */
  public function createXlsx(string $destination)
  {
    $this->getSharedStrings()->save();
    $zip = new PclZip($destination);
    $zip->create($this->unpackedDir, null, $this->unpackedDir);
  }
  
  /**
   * Переименовать лист
   *
   * @param int $number Идентификатор листа
   * @param string $newName Новое название листа
   * @return void
   * @throws Exception
   */
  public function renameWorksheet(int $number, string $newName)
  {
    if (!$this->workbookPath || !file_exists($this->workbookPath)) {
      throw new \Exception('Не найден файл workbook.xml');
    }
    $workbookXml = file_get_contents($this->workbookPath);
    $workbookXml = preg_replace('/(<sheet name=\")([^\"]+)(\"[^>]+sheetId=\"' . $number . '\"[^>]+>)/', '$1' . preg_quote($newName, '/') . '$3', $workbookXml);
    file_put_contents($this->workbookPath, $workbookXml);
  }
  
  /**
   * Удалить временные файлы
   *
   * @return void
   */
  public function clean()
  {
    self::rrmdir($this->unpackedDir);
  }
  
  /**
   * Рекурсивное удаление папки
   *
   * @param string $dirPath Путь к папке
   * @return void
   */
  protected static function rrmdir(string $dirPath)
  {
    if (is_dir($dirPath)) {
      $objects = scandir($dirPath);
      foreach ($objects as $object) {
        if ($object != "." && $object != "..") {
          if (is_dir($dirPath . DIRECTORY_SEPARATOR . $object) && !is_link($dirPath . "/" . $object))
            self::rrmdir($dirPath . DIRECTORY_SEPARATOR . $object);
          else
            unlink($dirPath . DIRECTORY_SEPARATOR . $object);
        }
      }
      rmdir($dirPath);
    }
  }
}
