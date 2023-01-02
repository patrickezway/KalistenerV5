/* %%DEBUT_ENTETE%%
-------------------------------------------------------------------------------
Fichier     : cm_phon.c
R{pertoire  : ./+libcm
Description :
-------------------------------------------------------------------------------
  Date   Auteur  Marque  Objet
../../..   ...    M_.    .
-------------------------------------------------------------------------------
%%FIN_ENTETE%% */

#include "../+INCLUDE/cm_public.h"


static int evoyelle( char );
static int econsonne( char );
static int supdoub( char * );
static int tra_A( char *, char *, int );
static int tra_B( char *, char *, int );
static int tra_C( char *, char *, int );
static int tra_D( char *, char *, int );
static int tra_E( char *, char *, int );
static int tra_G( char *, char *, int );
static int tra_I( char *, char *, int );
static int tra_L( char *, char *, int );
static int tra_M( char *, char *, int );
static int tra_N( char *, char *, int );
static int tra_O( char *, char *, int );
static int tra_P( char *, char *, int );
static int tra_Q( char *, char *, int );
static int tra_U( char *, char *, int );
static int tra_S( char *, char *, int );
static int tra_T( char *, char *, int );
static int tra_Y( char *, char *, int );
static int tra_FVWXZJK( char *, char *, int );
static void moteur( char * );


static int evoyelle( char c )

{
     if ( ( c == 65 ) || ( c == 69 ) || ( c == 73 ) || ( c == 79 ) ||
          ( c == 85 ) || ( c == 89 ) ) return( 1 );
     return( 0 );
}

static int econsonne( char c )

{
     if ( ( c != 65 ) && ( c != 69 ) && ( c != 73 ) && ( c != 79 ) &&
          ( c != 85 ) && ( c != 89 ) && ( c > 64 ) && ( c < 91 ) ) return( 1 );
     return( 0 );
}


static int supdoub( char *s )
{

     char tempo[80], c;
     int i, j;

     c = 0;
     for ( i=j=0; i < strlen( s ); i++)
     {
          if ( s[i] != c ) tempo[j++] = s[i];
          c = s[i];
     }
     tempo[j] = 0;
     strcpy( s, tempo);
     return( 1 );
}


void cm_conssupder( char *s )
{

     int i;

     for (i=strlen(s)-1; i > 0; i--)
     {
          if ( ( econsonne(s[i]-32) ) || ( s[i] == 'F' ) || (s[i] == 'J' ) )
          {
               s[i] = 0;
          }
          else
          {
               break;
          }
     }
}


void cm_letsupder( char *s )
{

     s[strlen(s)-1] = 0;
}


static int tra_A( char *s, char *tempo, int i )
{

     if ( s[i+1] == 'Y' )
     {
          if ( evoyelle(s[i+2]) == 0 )
          {
               strcat( tempo, "e");
               return( 2 );
          }
          else
          {
               strcat( tempo, "aJ");
               return( 2 );
          }
     }
     else
     if ( s[i+1] == 'I' )
     {
          if (((s[i+2] == 'N') || (s[i+2] == 'M')) && (evoyelle(s[i+3])==0))
          {
               strcat( tempo, "I");
               return( 3 );
          }
          else
          {
               if ( ( s[i+2] == 'L' ) && ( s[i+3] == 'L' ) )
               {
                    strcat( tempo, "aJ");
                    return( 4 );
               }
               strcat( tempo, "e");
               return( 2 );
          }
     }
     else
     if ( s[i+1] == 'U' )
     {
          strcat( tempo, "o");
          return( 2 );
     }
     else
     if ( ( s[i+1] == 'M' ) || ( s[i+1] == 'N' ) )
     {
          if ( ( evoyelle(s[i+2]) ) || ( s[i+1] == s[i+2] ) )
          {
               strcat( tempo, "a");
               return( 1 );
          }
          else
          {
               strcat( tempo, "A");
               return( 2 );
          }
     }
     strcat( tempo, "a");
     return( 1 );
}


static int tra_B( char *s, char *tempo, int i )
{

     if ( s[i+1] == 'B' )
     {
          strcat( tempo, "b");
          return( 2 );
     }
     strcat( tempo, "b");
     return( 1 );
}


static int tra_C( char *s, char *tempo, int i )
{

     if ( s[i+1] == 'A' )
     {
          strcat( tempo, "k");
          return( 1 );
     }
     else
     if ( s[i+1] == 'O' )
     {
          strcat( tempo, "k");
          return( 1 );
     }
     else
     if ( s[i+1] == 'U' )
     {
          strcat( tempo, "k");
          return( 1 );
     }
     else
     if ( s[i+1] == 'E' )
     {
          strcat( tempo, "s");
          return( 1 );
     }
     else
     if ( s[i+1] == 'I' )
     {
          strcat( tempo, "s");
          return( 1 );
     }
     else
     if ( s[i+1] == 'C' )
     {
          strcat( tempo, "k");
          if ( s[i+2] == 'H' ) return( 3 );
               else            return( 2 );
     }
     else
     if ( s[i+1] == 'E' )
     {
          strcat( tempo, "s");
          return( 1 );
     }
     if ( s[i+1] == 'H' )
     {
          if ( ( econsonne( s[i+2] ) ) && ( s[i+2] != 'M' ) )
          {
               strcat( tempo, "k");
               return( 2 );
          }
          else
          if ( evoyelle( s[i+2] ) )
          {
               if ( ( s[i+2] == 'A' ) && ( s[i+3] == 'O' ) )
               {
                    strcat( tempo, "k");
                    return( 2 );
               }
          }
          strcat( tempo, "F");
          return( 2 );
     }
     if ( evoyelle( s[i+1] ) )
     {
          strcat( tempo, "s");
          return( 1 );
     }
     strcat( tempo, "k");
     return( 1 );
}


static int tra_D( char *s, char *tempo, int i )
{

     if ( s[i+1] == 'D' )
     {
          strcat( tempo, "d");
          return( 2 );
     }
     strcat( tempo, "d");
     return( 1 );
}


static int tra_E( char *s, char *tempo, int i )
{

     if ( s[i+1] == 'A' )
     {
          if ( s[i+2] == 'I' )
          {
               strcat( tempo, "e");
               return( 3 );
          }
          else
          if ((s[i+2] == 'N')&&(evoyelle(s[i+3])==0)&&(econsonne(s[i+3])==0))
          {
               strcat( tempo, "A");
               return( 3 );
          }
          else
          if ( s[i+2] == 'U' )
          {
               strcat( tempo, "o");
               return( 3 );
          }
          else
          {
               strcat( tempo, "a");
               return( 2 );
          }
     }
     else
     if ( s[i+1] == 'M' )
     {
          if ( s[i+2] == 'M' )
          {
               strcat( tempo, "a");
               return( 3 );
          }
          else
          if ( evoyelle(s[i+2] ) )
          {
               strcat( tempo, "e");
               return( 1 );
          }
          else
          {
               strcat( tempo, "A");
               return( 2 );
          }
     }
     else
     if ( s[i+1] == 'N' )
     {
          if ( s[i+2] == 'N' )
          {
               strcat( tempo, "e");
               return( 1 );
          }
          else
          if ( evoyelle( s[i+2]) )
          {
               strcat( tempo, "e");
               return( 1 );
          }
          else
          {
               strcat( tempo, "A");
               return( 2 );
          }
     }
     else
     if ( ( s[i+1] == 'I' ) || ( s[i+1] == 'Y' ) )
     {
          if ( s[i+2] == 'N' )
          {
               strcat( tempo, "I");
               return( 3 );
          }
          else
          {
               strcat( tempo, "e");
               return( 2 );
          }
     }
     else
     if ( s[i+1] == 'U' )
     {
          strcat( tempo, "E");
          return( 2 );
     }
     if ( ( econsonne(s[i+1]) ) || ( evoyelle(s[i+1]) ) )
     {
          strcat( tempo, "e");
          return( 1 );
     }
     strcat( tempo, "e");
     return( 1 );
}


static int tra_G( char *s, char *tempo, int i )
{

     if ( ( s[i+1] == 'A' ) || ( s[i+1] == 'O' ) )
     {
          strcat( tempo, "g");
          return( 1 );
     }
     else
     if ( s[i+1] == 'U' )
     {
          strcat( tempo, "g");
          return( 2 );
     }
     else
     if ( ( s[i+1] == 'I' ) || ( s[i+1] == 'E' ) )
     {
          strcat( tempo, "j");
          return( 1 );
     }
     strcat( tempo, "g");
     return( 1 );
}


static int tra_I( char *s, char *tempo, int i )
{

     if ( ( s[i+1] == 'E' ) && ( s[i+2] == 'N' ) )
     {
          if ( evoyelle(s[i+3]) )
          {
               strcat( tempo, "i");
               return( 1 );
          }
          else
          {
               strcat( tempo, "iI");
               return( 3 );
          }
     }
     else
     if ( ( s[i+1] == 'L' ) && ( s[i+2] == 'L' ) )
     {
          if ( ( i != 0 ) && ( evoyelle(s[i-1]) ) )
          {
               strcat( tempo, "J");
               return( 3 );
          }
          else
          {
               strcat( tempo, "iJ");
               return( 3 );
          }
     }
     else
     if ( ( s[i+1] == 'E' ) && ( s[i+2] == 'R' ) )
     {
          if ( evoyelle(s[i+3]) == 0 )
          {
               strcat( tempo, "Je");
               return( 3 );
          }
     }
     else
     if ((s[i+1] == 'E') && (econsonne(s[i+2]) == 0) && (evoyelle(s[i+2])==0))
     {
          strcat( tempo, "i");
          return( 2 );
     }
     else
     if ( evoyelle(s[i+1]) )
     {
          strcat( tempo, "J");
          return( 1 );
     }
     else
     if ( s[i+1] == 'M' )
     {
          if ( s[i+2] == 'M' )
          {
               strcat( tempo, "i");
               return( 1 );
          }
          else
          if ( evoyelle(s[i+2] ) )
          {
               strcat( tempo, "i");
               return( 1 );
          }
          else
          {
               strcat( tempo, "I");
               return( 2 );
          }
     }
     else
     if ( s[i+1] == 'N' )
     {
          if ( s[i+2] == 'N' )
          {
               strcat( tempo, "i");
               return( 1 );
          }
          else
          if ( evoyelle( s[i+2]) )
          {
               strcat( tempo, "i");
               return( 1 );
          }
          else
          {
               strcat( tempo, "I");
               return( 2 );
          }
     }
     strcat( tempo, "i");
     return( 1 );
}


static int tra_L( char *s, char *tempo, int i )
{

     if ( s[i+1] == 'L' )
     {
          strcat( tempo, "l");
          return( 2 );
     }
     strcat( tempo, "l");
     return( 1 );
}


static int tra_M( char *s, char *tempo, int i )
{

     if ( s[i+1] == 'M' )
     {
          strcat( tempo, "m");
          return( 2 );
     }
     strcat( tempo, "m");
     return( 1 );
}


static int tra_N( char *s, char *tempo, int i )
{

     if ( s[i+1] == 'N' )
     {
          strcat( tempo, "n");
          return( 2 );
     }
     strcat( tempo, "n");
     return( 1 );
}


static int tra_O( char *s, char *tempo, int i )
{

     if ( s[i+1] == 'I' )
     {
          if ( s[i+2] == 'N' )
          {
               strcat( tempo, "oI");
               return( 3 );
          }
          else
          {
               strcat( tempo, "wa");
               return( 2 );
          }
     }
     else
     if ( s[i+1] == 'Y' )
     {
          strcat( tempo, "wa");
          return( 2 );
     }
     else
     if ( s[i+1] == 'U' )
     {
          if ( ( evoyelle(s[i+2])) && ( ( s[i+2] != 'I') || ( s[i+3] != 'L')))
          {
               strcat( tempo, "w");
               return( 2 );
          }
          else
          {
               strcat( tempo, "u");
               return( 2 );
          }
     }
     else
     if ( s[i+1] == 'M' )
     {
          if ( ( evoyelle(s[i+2]) ) || ( s[i+2] == 'M' ) )
          {
               strcat( tempo, "o");
               return( 1 );
          }
          else
          {
               strcat( tempo, "O");
               return( 2 );
          }
     }
     else
     if ( s[i+1] == 'N' )
     {
          if ( ( evoyelle(s[i+2]) ) || ( s[i+2] == 'N' ) )
          {
               strcat( tempo, "o");
               return( 1 );
          }
          else
          {
               strcat( tempo, "O");
               return( 2 );
          }
     }
     else
     if ( ( s[i+1] == 'E' ) && ( s[i+2] == 'U' ) )
     {
          strcat( tempo, "E");
          return( 3 );
     }
     strcat( tempo, "o");
     return( 1 );
}


static int tra_P( char *s, char *tempo, int i )
{

     if ( s[i+1] == 'P' )
     {
          strcat( tempo, "p");
          return( 2 );
     }
     else
     if ( s[i+1] == 'H' )
     {
          strcat( tempo, "f");
          return( 2 );
     }
     strcat( tempo, "p");
     return( 1 );
}


static int tra_Q( char *s, char *tempo, int i )
{

     if ( s[i+1] == 'U' )
     {
          strcat( tempo, "k");
          return( 2 );
     }
     strcat( tempo, "k");
     return( 1 );
}


static int tra_U( char *s, char *tempo, int i )
{

     if ( s[i+1] == 'M' )
     {
          if ( ( evoyelle(s[i+2]) ) || ( s[i+2] == 'M' ) )
          {
               strcat( tempo, "y");
               return( 1 );
          }
          else
          {
               strcat( tempo, "I");
               return( 2 );
          }
     }
     else
     if ( s[i+1] == 'N' )
     {
          if ( ( evoyelle(s[i+2]) ) || ( s[i+2] == 'N' ) )
          {
               strcat( tempo, "y");
               return( 1 );
          }
          else
          {
               strcat( tempo, "I");
               return( 2 );
          }
     }
     else
     strcat( tempo, "y");
     return( 1 );
}


static int tra_R( char *s, char *tempo, int i )
{

     if ( s[i+1] == 'R' )
     {
          strcat( tempo, "r");
          return( 2 );
     }
     strcat( tempo, "r");
     return( 1 );
}


static int tra_S( char *s, char *tempo, int i )
{

     if ( s[i+1] == 'S' )
     {
          strcat( tempo, "s");
          return( 2 );
     }
     else
     if ( s[i+1] == 'H' )
     {
          strcat( tempo, "F");
          return( 2 );
     }
     else
     if ( s[i+1] == 'C' )
     {
          if ( s[i+2] == 'H' )
          {
               strcat( tempo, "F");
               return( 3 );
          }
          else
          {
               strcat( tempo, "s");
               return( 2 );
          }
     }
     else
     if ( ( i != 0 ) && ( evoyelle( s[i-1] ) ) && ( evoyelle( s[i+1] ) ) )
     {
          strcat( tempo, "z");
          return( 1 );
     }
     strcat( tempo, "s");
     return( 1 );
}


static int tra_T( char *s, char *tempo, int i )
{

     if ( s[i+1] == 'T' )
     {
          strcat( tempo, "t");
          return( 2 );
     }
     else
     if ( s[i+1] == 'H' )
     {
          strcat( tempo, "t");
          return( 2 );
     }
     else
     if ( ( i != 0 ) && ( s[i-1] == 'A' ) && ( s[i+1] == 'I' ) &&
          ( s[i+2] == 'E' ) && ( s[i+3] == 'N' ) )
     {
          strcat( tempo, "s");
          return( 1 );
     }
     strcat( tempo, "t");
     return( 1 );
}


static int tra_Y( char *s, char *tempo, int i )
{

     if ( ( s[i+1] == 'M' ) || ( s[i+1] == 'N' ) )
     {
          strcat( tempo, "I");
          return( 2 );
     }
     else
     if ( ( i != 0 ) && ( evoyelle(s[i-1]) ) && ( evoyelle(s[i+1]) ) )
     {
          strcat( tempo, "J");
          return( 1 );
     }
     strcat( tempo, "i");
     return( 1 );
}


static int tra_FVWXZJK( char *s, char *tempo, int i )
{

     if ( s[i] == 'F' )
     {
          strcat( tempo, "f");
          return( 1 );
     }
     else
     if ( s[i] == 'V' )
     {
          strcat( tempo, "v");
          return( 1 );
     }
     else
     if ( s[i] == 'W' )
     {
          strcat( tempo, "v");
          return( 1 );
     }
     else
     if ( s[i] == 'X' )
     {
          strcat( tempo, "ks");
          return( 1 );
     }
     else
     if ( s[i] == 'Z' )
     {
          strcat( tempo, "z");
          return( 1 );
     }
     else
     if ( s[i] == 'J' )
     {
          strcat( tempo, "j");
          return( 1 );
     }
     else
     if ( s[i] == 'K' )
     {
          strcat( tempo, "k");
          return( 1 );
     }
}


static void moteur( char *s )
{

     int i;
     char tempo[80];

     for ( i=0; i < 80; tempo[i++]=0);
     for ( i=0; i < strlen(s);)
     {
          if ( s[i] == 'A' ) i += tra_A( s, tempo, i);
          else
          if ( s[i] == 'B' ) i += tra_B( s, tempo, i);
          else
          if ( s[i] == 'C' ) i += tra_C( s, tempo, i);
          else
          if ( s[i] == 'D' ) i += tra_D( s, tempo, i);
          else
          if ( s[i] == 'E' ) i += tra_E( s, tempo, i);
          else
          if ( s[i] == 'G' ) i += tra_G( s, tempo, i);
          else
          if ( s[i] == 'H' ) i++;
          else
          if ( s[i] == 'I' ) i += tra_I( s, tempo, i);
          else
          if ( s[i] == 'L' ) i += tra_L( s, tempo, i);
          else
          if ( s[i] == 'M' ) i += tra_M( s, tempo, i);
          else
          if ( s[i] == 'N' ) i += tra_N( s, tempo, i);
          else
          if ( s[i] == 'O' ) i += tra_O( s, tempo, i);
          else
          if ( s[i] == 'P' ) i += tra_P( s, tempo, i);
          else
          if ( s[i] == 'Q' ) i += tra_Q( s, tempo, i);
          else
          if ( s[i] == 'U' ) i += tra_U( s, tempo, i);
          else
          if ( s[i] == 'R' ) i += tra_R( s, tempo, i);
          else
          if ( s[i] == 'S' ) i += tra_S( s, tempo, i);
          else
          if ( s[i] == 'T' ) i += tra_T( s, tempo, i);
          else
          if ( s[i] == 'Y' ) i += tra_Y( s, tempo, i);
          else
          if ( ( s[i] == 'F' ) || ( s[i] == 'V' ) || ( s[i] == 'W' ) ||
               ( s[i] == 'X' ) || ( s[i] == 'Z' ) || ( s[i] == 'J' ) ||
               ( s[i] == 'K' ) )
          {
               i += tra_FVWXZJK( s, tempo, i);
          }
          else
          {
               i++;
               cm_conssupder( tempo );
          }
     }
     supdoub( tempo );
     cm_conssupder( tempo );
     strcpy( s, tempo);
}


void cm_phon( char *nom, char *phon )
{

     strcpy( phon, nom);
     moteur( phon );
     if( strlen(phon) == 0 )  strcpy( phon, "     " );
}
