foreach($goodsList as $v){
  sleep(1);
  $url =  $this->sign($v['product_code']);
  if($this->analyzeXml($url,$v['product_id'])){
      $sum++;
      echo "成功商品id:{$v['product_id']}";
      echo '<br/>';
      flush();
  }
  ob_flush();
}
