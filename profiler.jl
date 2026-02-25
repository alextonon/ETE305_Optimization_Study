using Profile

Profile.clear()        # nettoie le profil précédent
@profile begin
    include("eod_annuel.jl")
end

using ProfileView
ProfileView.view()     # ouvre une interface graphique pour visualiser le profil