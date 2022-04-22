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
    $workbookXml = preg_replace('/(<sheet name=\")([^\"]+)(\"[^>]+sheetId=\"' . $number . '\"[^>]+>)/', '${1}' . $newName. '${3}', $workbookXml);
    file_put_contents($this->workbookPath, $workbookXml);
  }
  
  /**
   * Клонирование листа. Клонирует любой существующий лист в конец документа с указанным названием
   *
   * @param int $number Номер листа, который собираемся клонировать
   * @param string $newName Новое название для нового листа
   * @return int Идентификатор нового листа
   * @return void
   * @throws Exception
   */
  public function cloneWorksheet(int $number, string $newName)
  {
    if (!$this->workbookPath || !file_exists($this->workbookPath)) {
      throw new \Exception('Не найден файл workbook.xml');
    }
    
    // Получаем номер нового листа (номер последнего +1)
    $matchesWorkbook = [];
    $workbookXml = file_get_contents($this->workbookPath);
    preg_match_all('/<sheet[^>]+sheetId=\"([^\"]+)\"[^>]+>/', $workbookXml, $matchesWorkbook);
    if (empty($matchesWorkbook) || empty($matchesWorkbook[1])) {
      throw new \Exception('Ошибка при поиске листов в workbook');
    }
    $newWorksheetNumber = (int)max($matchesWorkbook[1]) + 1;
    
    if ($newWorksheetNumber < 2) {
      throw new \Exception('Ошибка при определении номера клонируемого листа');
    }
    
    // Определяем идентификатор для листа
    $relsXml = [];
    $workbookRelsXml = file_get_contents($this->unpackedDir . '/xl/_rels/workbook.xml.rels');
    preg_match_all('/<Relationship[^>]Id=\"rId([0-9]+)\"/', $workbookRelsXml, $relsXml);
    if (empty($relsXml) || empty($relsXml[1])) {
      throw new \Exception('Ошибка при поиске rId в workbook.xml.rels');
    }
    $newRId = (int)max($relsXml[1]) + 1;
    if ($newRId < 2) {
      throw new \Exception('Ошибка при определении номера клонируемого листа');
    }
    
    // Добавляем запись с информацией о новом листе в файл со ссылками (workbook.xml.rels)
    $workbookRelsXml = str_replace(
      '</Relationships>',
      '<Relationship Id="rId' . $newRId . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' . $newWorksheetNumber . '.xml"/></Relationships>',
      $workbookRelsXml
    );
    file_put_contents($this->unpackedDir . '/xl/_rels/workbook.xml.rels', $workbookRelsXml);
    
    // Копируем содержимое листа
    if (!copy($this->unpackedDir . '/xl/worksheets/sheet' . $number . '.xml', $this->unpackedDir . '/xl/worksheets/sheet' . $newWorksheetNumber . '.xml')) {
      throw new \Exception('Ошибка при копировании worksheet.xml');
    }
    
    // Добавляем запись с информацией о новом листе в файл workbook.xml
    $workbookXml = str_replace(
      '</sheets>',
      '<sheet name="' . $newName . '" sheetId="' . $newWorksheetNumber . '" r:id="rId' . $newRId . '"/></sheets>',
      $workbookXml
    );
  
    file_put_contents($this->workbookPath, $workbookXml);
  
    return $newWorksheetNumber;
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
