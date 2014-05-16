#! perl
#
# putty-shortcut-sync.pl
#    read PuTTY sessions from the registry
#    create shortcuts to each of those sessions in a folder
#    (you could, for example, put this folder in your start menu)
#
# %YOU%\Local Data\Application Support\putty\
#   putty-hosts-toolbar     - shortcuts will be createdin here
#   icons           - images named after hostname/putty session names
#   putty.exe       - will be looked for in your path, then here...
#
# Author: f00li5h,  Licence: Same as perl.


use warnings; use strict;
use Data::Dumper;

use File::HomeDir;
use File::Spec::Functions qw[ catdir catfile splitdir ];

# put the shortcuts here (named after their putty session name)
my $SHORTCUT_DIR  = catdir (
    File::HomeDir->my_data,
    "putty", "putty-hosts-toolbar"
);

# find icon files in here when you create a shortcut
my $ICONS_IN  = catdir (
    File::HomeDir->my_data,
    "putty", "icons"
);

# set icons on the shortcuts - will help you know what you're clicking on
my $RANDOM_ICONS  = 0;
#     <true>        only if there isn't one
#     all        replace the icon with a random one

# putty is here, this is put in the shortcut
my $PATH_TO_PUTTY = (map {
            my $path = catfile ($_,'putty.exe');
                   -e $path ? $path : ()
                } split /;/, $ENV{PATH} )[0]
        || catfile (
            File::HomeDir->my_data,
            "putty", "putty.exe"
        );

# spew useless crap to stderr
my $DEBUG = 0;

# help!
my $help = 0;

# usage message, including the defaults
my $message =
"$0 [options] [session names]

Options [default]:
 --verbose   spew debug messages to STDERR [$DEBUG]
 --target    create shortcuts here
             [$SHORTCUT_DIR]
 --putty     the path to putty
             [$PATH_TO_PUTTY]
 --icons     look in here for images (named either the session or host name)
         [$ICONS_IN]
 --random    use random icons for a session if there is
         not an icon in --icons
         [$RANDOM_ICONS]

Session Names:
 leaving the name list emtpy means do all of them
 sessions don't need to exist in PuTTY yet, but the shortcut won't do anything helpful until you create a session with that name.
";

use Getopt::Long;
my  $opts_ok = GetOptions (my %getop=(
    'random'         => \$RANDOM_ICONS,
    'verbose|debug'  => \$DEBUG,            'help|?'  => \$help,
    'target=s'       => \$SHORTCUT_DIR,     'putty=s' => \$PATH_TO_PUTTY,
    'icons=s'        => sub { warn "--icons=$ICONS_IN does not exist\n" if !-d ($ICONS_IN = $_[1]) },
));

warn Dumper(\%getop) if $DEBUG;

# really need putty,
warn "--putty=$PATH_TO_PUTTY does not exist\n" and $opts_ok = 0 unless -f $PATH_TO_PUTTY;

# optional, so non-fatal
warn "--target=$SHORTCUT_DIR does not exist (I'll create it)\n" unless -d $SHORTCUT_DIR;

die $message if $help or not $opts_ok;

use constant
    PUTTY_REG_KEY => 'HKEY_CURRENT_USER\\Software\\SimonTatham\\PuTTY\\Sessions';


# we're going to snag your PuTTY sessions from the registry
my %registry;
use Win32::TieRegistry
    Delimiter    => "\\",
    TiedHash     => \%registry,
;

# take a list of session names to update shortcuts for
my @session_names = @ARGV;

# default to updating everything...
if (not @session_names ) {
    @session_names = map { s/\\+$//g; $_ }
             keys %{ $registry{ PUTTY_REG_KEY() } };

    warn "doing all sessions: ",(join ', ',  @session_names),"\n"
            if $DEBUG;
}

# create the specified directory if it doesn't exist
if (! -d $SHORTCUT_DIR) {
    my @path_parts = splitdir($SHORTCUT_DIR);
    warn Dumper \@path_parts ;
    my $path = shift @path_parts;
    while( @path_parts and $path = catdir($path, shift @path_parts) ){
        warn "creating $path";
        mkdir $path
    }
}

# read the whole icondir ahead of time, filter it later.
my @icondir_contents;
if (-d $ICONS_IN ) {
    opendir my($icondir), $ICONS_IN;
    @icondir_contents = grep {$_ ne '..' and $_ ne '.'} readdir($icondir);
}

# things in PuTTY session names that you're unlikely to want to use as shortcuts
my @silly_names = ('Default Settings', qw[ . .. ]);

# let's make some shortcuts.
use Win32::Shortcut;
for my $session_key (@session_names) {

    # session names contain some kind of crazy entity escaping
    (my $session = $session_key) =~ s/%([0-9a-z]{2})/chr(hex "$1")/gie;

    # that's a dumb shortcut to add
    if (grep $_ eq $session, @silly_names) {;
        warn "'$session' is a silly name for a shortcut, so I'll skip it.\n" if $DEBUG;
        next;
    }

    my $record = $registry{ PUTTY_REG_KEY() . '\\' . $session_key };

    # Comment: SESSION - host [forwards],
    my $desc = ( defined $record->{HostName}
            # drop SESSION if it's the start of the host name
            and ($record->{HostName}||'') !~ /^\Q$session/
                ? "$session " . $record->{HostName}
                :  $record->{HostName})
         . ($record->{PortForwardings} ? " [has forwards]" : '')
         ;

    my $link     = Win32::Shortcut->new();
    my $filename =  "$session.lnk";

    # this will update shortcuts, which seems like a sensible thing to do if they exist
    $link->Load(catfile($SHORTCUT_DIR , $filename ));
    $link->{Description} = $desc;
    $link->{Path       } = $PATH_TO_PUTTY;
    $link->{Arguments  } = "-load $session";

    # set an icon for this shortcut.
    if (-d $ICONS_IN and not $RANDOM_ICONS ) {

        # explorer lets you use some image formats, .bmps, .ico, and .dll here
        my @icons =
            map { catdir($ICONS_IN,$_) }

            grep {
            # find by hostname
            (defined $record->{HostName}
                    and $record->{HostName} ne ''
                        ?/^\Q($record->{'HostName'})/ : 0 )
            # then try session name
            or (defined $session
                    and $session ne ''
                        ? /^\Q$session/               : 0 )
             } @icondir_contents;

        # if we found something, set it as the icon
        if (@icons and -f (my $icon_name = shift @icons) ) {
            $link->{IconLocation} = $icon_name;
            $link->{'IconNumber'} = 0 ; # the first one is fine.
            warn "using $icon_name/0 as the icon for $session"
                if $DEBUG;
        }
    }

    # do some random icons.
    if (    ( $RANDOM_ICONS eq 'all')
         or ( $RANDOM_ICONS and $link->{IconLocation} eq '' )
      ) {
        $link->{IconLocation} = "%SystemRoot%\\system32\\SHELL32.dll";
        # there were this many on XP...
        $link->{'IconNumber'} = int rand 237;
    }

    # and save it all off
    $link->Save(catfile($SHORTCUT_DIR , $filename ));
    warn Dumper( $link ) if $DEBUG;
    $link->Close();
}

# NB: I updated 1 shortcut(s) is bullshit, you lazy bastard.
print "I updated ", scalar @session_names , " shortcut" ,
      1==@session_names ?'':'s',
      " in ", $SHORTCUT_DIR, ": ", join ', ', @session_names
