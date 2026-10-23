<?php
/**
 * As configurações básicas do WordPress
 *
 * O script de criação wp-config.php usa esse arquivo durante a instalação.
 * Você não precisa usar o site, você pode copiar este arquivo
 * para "wp-config.php" e preencher os valores.
 *
 * Este arquivo contém as seguintes configurações:
 *
 * * Configurações do banco de dados
 * * Chaves secretas
 * * Prefixo do banco de dados
 * * ABSPATH
 *
 * @link https://wordpress.org/documentation/article/editing-wp-config-php/
 *
 * @package WordPress
 */

// ** Configurações do banco de dados - Você pode pegar estas informações com o serviço de hospedagem ** //
/** O nome do banco de dados do WordPress */
define( 'DB_NAME', 'if0_43238904_doom' );

/** Usuário do banco de dados MySQL */
define( 'DB_USER', 'if0_43238904' );

/** Senha do banco de dados MySQL */
define( 'DB_PASSWORD', '9UPUBoUcJZ' );

/** Nome do host do MySQL */
define( 'DB_HOST', 'sql300.infinityfree.com' );

/** Charset do banco de dados a ser usado na criação das tabelas. */
define( 'DB_CHARSET', 'utf8mb4' );

/** O tipo de Collate do banco de dados. Não altere isso se tiver dúvidas. */
define( 'DB_COLLATE', '' );

/**#@+
 * Chaves únicas de autenticação e salts.
 *
 * Altere cada chave para um frase única!
 * Você pode gerá-las
 * usando o {@link https://api.wordpress.org/secret-key/1.1/salt/ WordPress.org
 * secret-key service}
 * Você pode alterá-las a qualquer momento para invalidar quaisquer
 * cookies existentes. Isto irá forçar todos os
 * usuários a fazerem login novamente.
 *
 * @since 2.6.0
 */
define( 'AUTH_KEY',         '%tY*PitrW&Wq0p+ZOdP%^5|g9[ahZ.ZGj)8qM471Rv@ze9uSh(VBzFNhw%:LWmLq' );
define( 'SECURE_AUTH_KEY',  '90P=9;&h(q@&s ug[tb 2|u@!0fvP5K[G:9n~<CYJ(~V: s7=7=-P3t,Zg{f16sv' );
define( 'LOGGED_IN_KEY',    's2qAB:@1g^;aP]Gr~y`|Y%^|RZ>~,<CvVRCDx+!b=0,!h:ZBNh`8iPdKXRN9F3 q' );
define( 'NONCE_KEY',        'PQ`}MF*%# HN%`~lwC?.+t9Ob_|5{Sp7zX:JoV-eAu{ <?F,cA6:EB?;6{.$~O5X' );
define( 'AUTH_SALT',        'Nekp~|CIN(7}%NZq}DEu^3:~4TKq.Z6YipRj3XM6R)(VU1$/CM&M&K0dK.{u:*t*' );
define( 'SECURE_AUTH_SALT', 'm4:EP*hf2m!EKUtcz>Ejhw/)MS[(5yHV,%QFtZy2l!-i-23|Oama&ySWGAnVjB2(' );
define( 'LOGGED_IN_SALT',   'O&-=}8g&v]nkoP+/pN3vd09w[bg|(:[*Aaw6GjG)+<L%mvTG#,@#>-TT *5PDdtB' );
define( 'NONCE_SALT',       'N.KYCe3uIv+(u&1-?|HR[e17CKb9/DI@u`mIOO,iA>y_R^87+YA3Z= 8@!O>$GUs' );

/**#@-*/

/**
 * Prefixo da tabela do banco de dados do WordPress.
 *
 * Você pode ter várias instalações em um único banco de dados se você der
 * um prefixo único para cada um. Somente números, letras e sublinhados!
 */
$table_prefix = 'wp_';

/**
 * Para desenvolvedores: Modo de debug do WordPress.
 *
 * Altere isto para true para ativar a exibição de avisos
 * durante o desenvolvimento. É altamente recomendável que os
 * desenvolvedores de plugins e temas usem o WP_DEBUG
 * em seus ambientes de desenvolvimento.
 *
 * Para informações sobre outras constantes que podem ser utilizadas
 * para depuração, visite o Codex.
 *
 * @link https://wordpress.org/documentation/article/debugging-in-wordpress/
 */
define( 'WP_DEBUG', false );

/* Adicione valores personalizados entre esta linha até "Isto é tudo". */



/* Isto é tudo, pode parar de editar! :) */

/** Caminho absoluto para o diretório WordPress. */
if ( ! defined( 'ABSPATH' ) ) {
	define( 'ABSPATH', __DIR__ . '/' );
}

/** Configura as variáveis e arquivos do WordPress. */
require_once ABSPATH . 'wp-settings.php';
