<?php
/**
 * Desc:
 * User: baagee
 * Date: 2020/6/11
 * Time: 下午2:39
 */

namespace BaAGee\Excel;

/**
 * Class WriterBuffer
 * @package BaAGee\Excel
 */
class WriterBuffer
{
    /**
     * @var bool|resource|null
     */
    protected $fd = null;
    /**
     * @var string
     */
    protected $buffer = '';
    /**
     * @var bool
     */
    protected $checkUtf8 = false;

    /**
     * WriterBuffer constructor.
     * @param        $filename
     * @param string $fdFopenFlags
     * @param bool   $checkUtf8
     */
    public function __construct($filename, $fdFopenFlags = 'w', $checkUtf8 = false)
    {
        $this->checkUtf8 = $checkUtf8;
        $this->fd = fopen($filename, $fdFopenFlags);
        if ($this->fd === false) {
            XLSXWriter::log("Unable to open $filename for writing.");
        }
    }

    /**
     * @param $string
     */
    public function write($string)
    {
        $this->buffer .= $string;
        if (isset($this->buffer[8191])) {
            $this->purge();
        }
    }

    /**
     *
     */
    protected function purge()
    {
        if ($this->fd) {
            if ($this->checkUtf8 && !self::isValidUTF8($this->buffer)) {
                XLSXWriter::log("Error, invalid UTF8 encoding detected.");
                $this->checkUtf8 = false;
            }
            fwrite($this->fd, $this->buffer);
            $this->buffer = '';
        }
    }

    /**
     *
     */
    public function close()
    {
        $this->purge();
        if ($this->fd) {
            fclose($this->fd);
            $this->fd = null;
        }
    }

    /**
     *
     */
    public function __destruct()
    {
        $this->close();
    }

    /**
     * @return bool|int
     */
    public function fTell()
    {
        if ($this->fd) {
            $this->purge();
            return ftell($this->fd);
        }
        return -1;
    }

    /**
     * @param $pos
     * @return int
     */
    public function fSeek($pos)
    {
        if ($this->fd) {
            $this->purge();
            return fseek($this->fd, $pos);
        }
        return -1;
    }

    /**
     * @param $string
     * @return bool
     */
    protected static function isValidUTF8($string)
    {
        if (function_exists('mb_check_encoding')) {
            return mb_check_encoding($string, 'UTF-8') ? true : false;
        }
        return preg_match("//u", $string) ? true : false;
    }
}
