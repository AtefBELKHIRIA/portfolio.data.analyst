---
layout: default
---

<header>
   <h1 class="header-title">{{ site.title }}</h1>
  <img src="{{ site.profile_image }}" alt="Photo de profil" class="profile">
   <!-- Ajout d'une classe "description" -->
  <p class="description">{{ site.description }}</p>
  
  <div class="contact">
    <h3>Contact</h3>
    <p>{{ site.contact.phone }}</p>
    <p><a href="mailto:{{ site.contact.email }}">{{ site.contact.email }}</a></p>
  </div>

  <div class="social">
    <h3>Profils</h3>
    <p>
       <a href="{{ site.social.linkedin }}" target="_blank">LinkedIn</a><br>
       <a href="{{ site.social.github }}" target="_blank">GitHub</a>
    </p>
  </div>
</header>


<section>
  {% capture markdown_content %}
  {% include_relative README.md %}
  {% endcapture %}
  {{ markdown_content | markdownify }}
</section>
